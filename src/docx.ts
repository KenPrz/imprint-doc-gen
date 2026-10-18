import { advancedCloner, inlinePictures } from './docx-advanced.ts';
import { standardImageResolver } from './image-binding.ts';
import { RenderError } from './errors.ts';
import { compileStory, renderStory, identityAttributes, type RenderBudget, type ImagePlan } from './docx-story.ts';
import { lookup, scalar, type Scope } from './scope.ts';
import { validateData } from './data.ts';
import { createRenderPackage } from './render-package.ts';
import type { OfficePackage } from './ooxml.ts';
import type { CompiledTemplate, RenderOptions } from './types.ts';
import { resolvePart, relsPath } from './package.ts';
import { ns, indexXml, escapeXml, parseXml, descendants, type Edit } from './xml.ts';

export async function compileDocx({ parts: original, mainPart }: OfficePackage): Promise<CompiledTemplate> {
  const storyParts = [mainPart];
  const mainRelationships = original.get(relsPath(mainPart));
  if (mainRelationships) for (const link of descendants(parseXml(mainRelationships.toString('utf8'), relsPath(mainPart)), ns.packageRelationship, 'Relationship')) {
    if (['header', 'footer', 'footnotes', 'endnotes', 'comments'].some(kind => link.getAttribute('Type') === `${ns.relationship}/${kind}`)) storyParts.push(resolvePart(mainPart, link.getAttribute('Target')!));
  }
  // Reserve identities in every XML part, including untouched notes or drawings.
  const reservedIdentities = new Set<number>();
  for (const [part, bytes] of original) if (part.endsWith('.xml')) {
    for (const { attribute, hex } of indexXml(bytes.toString('utf8'), part).flatMap(identityAttributes)) {
      const value = Number.parseInt(attribute.value, hex ? 16 : 10);
      if (Number.isFinite(value)) reservedIdentities.add(value);
    }
  }
  let nextInlineIdentity=1;
  const inlineIdentity=()=>{while(reservedIdentities.has(nextInlineIdentity))nextInlineIdentity++;const id=nextInlineIdentity++;reservedIdentities.add(id);return String(id);};
  const plans = [...new Set(storyParts)].map(part => compileStory(part, inlinePictures(original.get(part)!.toString('utf8'),inlineIdentity)));
  for (const plan of plans) if (plan.images.length) {
    const path = relsPath(plan.part), bytes = original.get(path);
    if (!bytes && plan.images.some(i=>i.reference.value !== 'imprintInlinePlaceholder')) throw new RenderError('INVALID_PACKAGE', 'Picture is missing its relationship part', plan.part);
    const links = descendants(parseXml(bytes?.toString('utf8') ?? `<Relationships xmlns="${ns.packageRelationship}"/>`, path), ns.packageRelationship, 'Relationship');
    for (const image of plan.images) if (image.reference.value !== 'imprintInlinePlaceholder' && !links.some(link => link.getAttribute('Id') === image.reference.value && link.getAttribute('Type') === `${ns.relationship}/image`)) throw new RenderError('INVALID_PACKAGE', 'Picture references an undefined image relationship', plan.part);
  }
  return Object.freeze({ format: 'docx' as const, async render(data: Record<string, unknown>, options: RenderOptions = {}) {
    validateData(data);
    const attachmentMap = Object.assign(Object.create(null), options.attachments ?? {}) as Record<string, Uint8Array>;
    const output = createRenderPackage(original, mainPart, attachmentMap, options.images);
    const imageKey = standardImageResolver(attachmentMap);
    const usedIdentities = new Set(reservedIdentities); let nextIdentity = 1;
    const freshIdentity = (hex: boolean) => {
      while (usedIdentities.has(nextIdentity)) nextIdentity++;
      const id = nextIdentity++; usedIdentities.add(id);
      return hex ? id.toString(16).toUpperCase().padStart(8, '0') : String(id);
    };
    const budget: RenderBudget = { iterations: 0 };
    const renderedParts = new Map<string,Buffer>();
    const renderImage = async (binding: ImagePlan, scope: Scope, part: string): Promise<Edit[]> => {
        // Literal attachment keys retain milestone-1 behavior at root scope.
        // Inside loops, image paths resolve from the current item or explicit ../ parent.
        if (binding.standard) {
          const key = imageKey(binding.path, scope, part);
          if (!key && binding.reference.value === 'imprintInlinePlaceholder') throw new RenderError('MISSING_ATTACHMENT','Inline pictures require an attachment',part);
          if (!key) return [{ start: binding.description.start, end: binding.description.end, value: 'Image placeholder' }];
          const id = await output.addImage(part, key, binding.size);
          return [{ start: binding.reference.start, end: binding.reference.end, value: id }, { start: binding.description.start, end: binding.description.end, value: 'Image attachment' }];
        }
        const key = scope.length === 1 && Object.hasOwn(attachmentMap, binding.path) ? binding.path : scalar(lookup(scope, binding.path, part), binding.path, part);
        const id = await output.addImage(part, key, binding.size);
        return [{ start: binding.reference.start, end: binding.reference.end, value: id }, { start: binding.description.start, end: binding.description.end, value: escapeXml(`Attachment ${key}`) }];
    };
    const cloner = advancedCloner(original, mainPart, options.restartNumbering ?? false, async (source,part,scope) => {
      const notePlan = compileStory(part,inlinePictures(source,()=>freshIdentity(false)));
      if (scope.length>1 && notePlan.images.length) throw new RenderError('UNSUPPORTED_TEMPLATE','Repeated note picture bindings are not supported',part);
      if(notePlan.nodes.some(n=>n.uri===ns.word && ['footnoteReference','endnoteReference','commentReference'].includes(n.local)))throw new RenderError('UNSUPPORTED_TEMPLATE','References inside note definitions are unsupported',part);
      return renderStory(notePlan,data,budget,(binding,scope)=>renderImage(binding,scope,part),freshIdentity,undefined,scope);
    });
    for (const plan of plans) {
      if (cloner.isDefinition(plan.part)) continue;
      const result = await renderStory(plan, data, budget, (binding,scope)=>renderImage(binding,scope,plan.part), freshIdentity, cloner.fragment);

      if (result !== plan.source) { parseXml(result, plan.part); renderedParts.set(plan.part, Buffer.from(result)); }
    }
    await cloner.finish(renderedParts,data);
    for (const [part,bytes] of renderedParts) output.setPart(part,bytes);
    return output.finish();
  } });
}
