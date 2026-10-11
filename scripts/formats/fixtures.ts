import { writePackage, type PackageParts } from '../../src/package.ts';
import { ns, escapeXml as esc } from '../../src/xml.ts';
import { formatNs as f } from '../../src/format-namespaces.ts';
import { mediaTypes } from '../../src/types.ts';
import { png } from '../corpus/common.ts';
export type ExtraFormat = 'pptx' | 'odt' | 'ods' | 'odp';
export const sample = (level: number) => ({
  title: `Example document ${level}`,
  customer: level % 3 ? 'Alex Example' : 'Zoë Núñez',
  amount: level + 0.5,
  items: Array.from({ length: level >= 91 ? 10 : 3 }, (_, n) => ({
    name: `Item ${n + 1}`,
    amount: (n + 1) * 5,
  })),
});
export const red = png([220, 40, 40]),
  blue = png([40, 80, 220]);
const rels = (links: string) => `<Relationships xmlns="${ns.packageRelationship}">${links}</Relationships>`;
const link = (id: string, type: string, target: string) =>
  `<Relationship Id="${id}" Type="${ns.relationship}/${type}" Target="${target}"/>`;
const group =
  '<p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
const mapping =
  'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"';
export async function fixture(format: ExtraFormat, level: number, expected = false) {
  const data = sample(level),
    parts: PackageParts = new Map(),
    put = (path: string, content: string | Buffer) => parts.set(path, Buffer.from(content));
  const value = (path: 'title' | 'customer' | 'amount') => (expected ? String(data[path]) : `{d.${path}}`),
    image = level >= 11;
  if (format === 'pptx') {
    const decl = `xmlns:p="${f.presentation}" xmlns:a="${ns.drawing}" xmlns:r="${ns.relationship}"`;
    const text = (value: string) =>
      level >= 21 && value.startsWith('{')
        ? `<a:r><a:rPr lang="en-US" sz="${1800 + level * 2}"/><a:t>${esc(value.slice(0, 4))}</a:t></a:r><a:r><a:t>${esc(value.slice(4))}</a:t></a:r>`
        : `<a:r><a:rPr lang="en-US" sz="${1800 + level * 2}"/><a:t>${esc(value)}</a:t></a:r>`;
    const shape = (content: string, id = 2, y = 400000) =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="400000" y="${y}"/><a:ext cx="7000000" cy="900000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p>${text(content)}</a:p></p:txBody></p:sp>`;
    const picture = `<p:pic><p:nvPicPr><p:cNvPr id="4" name="Logo" descr="${expected ? 'Image attachment' : '{d.logo}'}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="image"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="8000000" y="400000"/><a:ext cx="1000000" cy="700000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
    const repeat = level >= 61,
      slides = repeat
        ? expected
          ? data.items.map((item) => item.name)
          : ['{d.items[i].name}', '{d.items[i+1].name}']
        : [value('customer')];
    let overrides = '';
    slides.forEach((title, n) => {
      put(
        `ppt/slides/slide${n + 1}.xml`,
        `<p:sld ${decl}><p:cSld><p:spTree>${group}${title.includes('[i+1]') ? '' : shape(value('title'))}${shape(title, 3, 1600000)}${image && !title.includes('[i+1]') ? picture : ''}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
      );
      put(
        `ppt/slides/_rels/slide${n + 1}.xml.rels`,
        rels(
          link('layout', 'slideLayout', '../slideLayouts/layout.xml') +
            (image ? link('image', 'image', '../media/logo.png') : ''),
        ),
      );
      overrides += `<Override PartName="/ppt/slides/slide${n + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
    });
    put(
      'ppt/presentation.xml',
      `<p:presentation ${decl}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="master"/></p:sldMasterIdLst><p:sldIdLst>${slides.map((_, n) => `<p:sldId id="${256 + n}" r:id="s${n}"/>`).join('')}</p:sldIdLst><p:sldSz cx="10000000" cy="5625000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
    );
    put(
      'ppt/_rels/presentation.xml.rels',
      rels(
        link('master', 'slideMaster', 'slideMasters/master.xml') +
          slides.map((_, n) => link(`s${n}`, 'slide', `slides/slide${n + 1}.xml`)).join(''),
      ),
    );
    put(
      'ppt/slideLayouts/layout.xml',
      `<p:sldLayout ${decl} type="blank"><p:cSld name="Blank"><p:spTree>${group}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
    );
    put(
      'ppt/slideLayouts/_rels/layout.xml.rels',
      rels(link('master', 'slideMaster', '../slideMasters/master.xml')),
    );
    put(
      'ppt/slideMasters/master.xml',
      `<p:sldMaster ${decl}><p:cSld><p:spTree>${group}</p:spTree></p:cSld><p:clrMap ${mapping}/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="layout"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,
    );
    put(
      'ppt/slideMasters/_rels/master.xml.rels',
      rels(
        link('layout', 'slideLayout', '../slideLayouts/layout.xml') +
          link('theme', 'theme', '../theme/theme.xml'),
      ),
    );
    put(
      'ppt/theme/theme.xml',
      `<a:theme xmlns:a="${ns.drawing}" name="Imprint"><a:themeElements><a:clrScheme name="Imprint">${['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].map((name, n) => `<a:${name}><a:srgbClr val="${n % 2 ? 'FFFFFF' : '164B61'}"/></a:${name}>`).join('')}</a:clrScheme><a:fontScheme name="Imprint"><a:majorFont><a:latin typeface="Liberation Sans"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Liberation Sans"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Imprint"><a:fillStyleLst>${'<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'.repeat(3)}</a:fillStyleLst><a:lnStyleLst>${'<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'.repeat(3)}</a:lnStyleLst><a:effectStyleLst>${'<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3)}</a:effectStyleLst><a:bgFillStyleLst>${'<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'.repeat(3)}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`,
    );
    put('_rels/.rels', rels(link('main', 'officeDocument', 'ppt/presentation.xml')));
    put(
      '[Content_Types].xml',
      `<Types xmlns="${ns.contentTypes}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/master.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/layout.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>${overrides}</Types>`,
    );
    if (image) put('ppt/media/logo.png', expected ? blue : red);
  } else {
    const decl =
      Object.entries(f)
        .filter(([name]) => name !== 'presentation')
        .map(([name, uri]) => `xmlns:${name}="${uri}"`)
        .join(' ') +
      ` xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"`;
    const text = (content: string) =>
      `<text:p>${level >= 21 && content.startsWith('{') ? esc(content.slice(0, 4)) + `<text:span>${esc(content.slice(4))}</text:span>` : esc(content)}</text:p>`;
    const picture = `<draw:frame draw:name="Logo" text:anchor-type="paragraph" svg:x="1cm" svg:y="1cm" svg:width="3cm" svg:height="1cm"><draw:image xlink:href="Pictures/logo.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/><svg:desc>${expected ? 'Image attachment' : '{d.logo}'}</svg:desc></draw:frame>`;
    const repeat = level >= 41,
      labels = repeat
        ? expected
          ? data.items.map((item) => item.name)
          : ['{d.items[i].name}', '{d.items[i+1].name}']
        : [value('customer')];
    const row = (label: string) =>
      `<table:table-row><table:table-cell office:value-type="string">${text(label)}</table:table-cell></table:table-row>`;
    const table = `<table:table table:name="Items"><table:table-column/>${labels.map(row).join('')}</table:table>`;
    let body =
      format === 'odt'
        ? `<office:text>${text(value('title'))}${image ? text('Logo') + picture : ''}${table}</office:text>`
        : format === 'ods'
          ? `<office:spreadsheet><table:table table:name="Items"><table:table-column/>${row(value('title'))}${labels.map(row).join('')}${image ? `<table:shapes>${picture}</table:shapes>` : ''}</table:table></office:spreadsheet>`
          : `<office:presentation>${labels.map((label, n) => `<draw:page draw:name="Slide${n + 1}" draw:master-page-name="Default"><draw:frame draw:name="Title${n}" svg:x="1cm" svg:y="2cm" svg:width="20cm" svg:height="8cm"><draw:text-box>${label.includes('[i+1]') ? '' : text(value('title'))}${text(label)}</draw:text-box></draw:frame>${image && !label.includes('[i+1]') ? picture : ''}</draw:page>`).join('')}</office:presentation>`;
    put('mimetype', mediaTypes[format]);
    put(
      'content.xml',
      `<office:document-content ${decl} office:version="1.3"><office:automatic-styles/><office:body>${body}</office:body></office:document-content>`,
    );
    put(
      'styles.xml',
      `<office:document-styles ${decl} office:version="1.3"><office:styles><style:default-style style:family="paragraph"><style:text-properties fo:font-size="${12 + (level % 5)}pt"/></style:default-style></office:styles><office:automatic-styles><style:page-layout style:name="page"><style:page-layout-properties fo:page-width="${format === 'odp' ? '28' : '21'}cm" fo:page-height="${format === 'odp' ? '15.75' : '29.7'}cm" fo:margin="1cm"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Default" style:page-layout-name="page"/></office:master-styles></office:document-styles>`,
    );
    if (image) put('Pictures/logo.png', expected ? blue : red);
    put(
      'META-INF/manifest.xml',
      `<manifest:manifest xmlns:manifest="${f.manifest}" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="${mediaTypes[format]}"/>${[
        ...parts.keys(),
      ]
        .filter((p) => p !== 'mimetype')
        .map(
          (path) =>
            `<manifest:file-entry manifest:full-path="${path}" manifest:media-type="${path.endsWith('.png') ? 'image/png' : 'text/xml'}"/>`,
        )
        .join('')}</manifest:manifest>`,
    );
  }
  return writePackage(parts);
}
