import { carboneImageResolver } from './image-binding.ts';
import { RenderError } from './errors.ts';
import { compileStory, renderStory, identityAttributes, type RenderBudget } from './docx-story.ts';
import { lookup, scalar } from './scope.ts';
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
    if (['header', 'footer'].some(kind => link.getAttribute('Type') === `${ns.relationship}/${kind}`)) storyParts.push(resolvePart(mainPart, link.getAttribute('Target')!));
  }
  const plans = [...new Set(storyParts)].map(part => compileStory(part, original.get(part)!.toString('utf8')));
  for (const plan of plans) if (plan.images.length) {
    const path = relsPath(plan.part), bytes = original.get(path);
    if (!bytes) throw new RenderError('INVALID_PACKAGE', 'Picture is missing its relationship part', plan.part);
    const links = descendants(parseXml(bytes.toString('utf8'), path), ns.packageRelationship, 'Relationship');
    for (const image of plan.images) if (!links.some(link => link.getAttribute('Id') === image.reference.value && link.getAttribute('Type') === `${ns.relationship}/image`)) throw new RenderError('INVALID_PACKAGE', 'Picture references an undefined image relationship', plan.part);
  }
  // Reserve identities in every XML part, including untouched notes or drawings.
  const reservedIdentities = new Set<number>();
  for (const [part, bytes] of original) if (part.endsWith('.xml')) {
    for (const { attribute, hex } of indexXml(bytes.toString('utf8'), part).flatMap(identityAttributes)) {
      const value = Number.parseInt(attribute.value, hex ? 16 : 10);
      if (Number.isFinite(value)) reservedIdentities.add(value);
    }
  }
  return Object.freeze({ format: 'docx' as const, async render(data: Record<string, unknown>, options: RenderOptions = {}) {
    validateData(data);
    const attachmentMap = Object.assign(Object.create(null), options.attachments ?? {}) as Record<string, Uint8Array>;
    const output = createRenderPackage(original, mainPart, attachmentMap);
    const imageKey = carboneImageResolver(attachmentMap);
    const usedIdentities = new Set(reservedIdentities); let nextIdentity = 1;
    const freshIdentity = (hex: boolean) => {
      while (usedIdentities.has(nextIdentity)) nextIdentity++;
      const id = nextIdentity++; usedIdentities.add(id);
      return hex ? id.toString(16).toUpperCase().padStart(8, '0') : String(id);
    };
    const budget: RenderBudget = { iterations: 0 };
    for (const plan of plans) {
      const result = await renderStory(plan, data, budget, async (binding, scope): Promise<Edit[]> => {
        // Literal attachment keys retain milestone-1 behavior at root scope.
        // Inside loops, image paths resolve from the current item or explicit ../ parent.
        if (binding.carbone) {
          const key = imageKey(binding.path, scope, plan.part);
          if (!key) return [{ start: binding.description.start, end: binding.description.end, value: 'Image placeholder' }];
          const id = await output.addImage(plan.part, key);
          return [{ start: binding.reference.start, end: binding.reference.end, value: id }, { start: binding.description.start, end: binding.description.end, value: 'Image attachment' }];
        }
        const key = scope.length === 1 && Object.hasOwn(attachmentMap, binding.path) ? binding.path : scalar(lookup(scope, binding.path, plan.part), binding.path, plan.part);
        const id = await output.addImage(plan.part, key);
        return [{ start: binding.reference.start, end: binding.reference.end, value: id }, { start: binding.description.start, end: binding.description.end, value: escapeXml(`Attachment ${key}`) }];
      }, freshIdentity);
      if (result !== plan.source) { parseXml(result, plan.part); output.setPart(plan.part, Buffer.from(result)); }
    }
    return output.finish();
  } });
}
