import { SaxesParser } from 'saxes';

/** โหนด XML แบบง่าย (ตรวจชื่อด้วย local name + namespace URI ไม่ขึ้นกับ prefix ในไฟล์) */
export interface XmlNode {
  local: string;
  uri: string;
  /** attribute ตาม local name (ไม่รวม xmlns) */
  attrs: Record<string, string>;
  /** attribute ตาม namespace URI + local name เช่น r:id */
  nsAttrs: Record<string, string>;
  children: XmlNode[];
  text: string;
}

export const XML_NS = {
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  main: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  pkgRels: 'http://schemas.openxmlformats.org/package/2006/relationships',
};

/** แปลง XML เป็นต้นไม้ (โยน error ถ้า XML ไม่ถูกรูปแบบ) */
export function parseXml(xml: string): XmlNode {
  const parser = new SaxesParser({ xmlns: true });
  const root: XmlNode = { local: '#root', uri: '', attrs: {}, nsAttrs: {}, children: [], text: '' };
  const stack: XmlNode[] = [root];
  parser.on('opentag', (tag) => {
    const node: XmlNode = { local: tag.local, uri: tag.uri, attrs: {}, nsAttrs: {}, children: [], text: '' };
    for (const attr of Object.values(tag.attributes)) {
      if (attr.prefix === 'xmlns' || attr.name === 'xmlns') continue;
      if (attr.uri) node.nsAttrs[`${attr.uri}|${attr.local}`] = attr.value;
      else node.attrs[attr.local] = attr.value;
    }
    stack[stack.length - 1].children.push(node);
    stack.push(node);
  });
  parser.on('text', (t) => {
    stack[stack.length - 1].text += t;
  });
  parser.on('cdata', (t) => {
    stack[stack.length - 1].text += t;
  });
  parser.on('closetag', () => {
    stack.pop();
  });
  parser.write(xml).close();
  const top = root.children[0];
  if (!top) throw new Error('XML ว่าง');
  return top;
}

export function child(node: XmlNode | undefined, uri: string, local: string): XmlNode | undefined {
  return node?.children.find((c) => c.uri === uri && c.local === local);
}

export function childrenOf(node: XmlNode | undefined, uri: string, local: string): XmlNode[] {
  return node ? node.children.filter((c) => c.uri === uri && c.local === local) : [];
}

/** หาโหนดลูกหลานทุกระดับ */
export function descendants(node: XmlNode, uri: string, local: string): XmlNode[] {
  const out: XmlNode[] = [];
  const walk = (n: XmlNode) => {
    for (const c of n.children) {
      if (c.uri === uri && c.local === local) out.push(c);
      walk(c);
    }
  };
  walk(node);
  return out;
}

/** ข้อความทั้งหมดใต้โหนด (ต่อกัน) */
export function textOf(node: XmlNode | undefined): string {
  if (!node) return '';
  return node.text + node.children.map(textOf).join('');
}

export function escapeXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    // อักขระควบคุมที่ XML 1.0 ไม่อนุญาต
    // eslint-disable-next-line no-control-regex
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
}
