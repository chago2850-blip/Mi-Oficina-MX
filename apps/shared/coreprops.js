// Lee las propiedades del documento (docProps/core.xml) de un archivo de Office. GPL-3.0
import JSZip from 'jszip';
export async function readCoreProps(bytes) {
  try {
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('docProps/core.xml')?.async('string');
    if (!xml) return {};
    const d = new DOMParser().parseFromString(xml, 'application/xml');
    const g = (tag) => d.getElementsByTagName(tag)[0]?.textContent?.trim() || '';
    return { title: g('dc:title'), author: g('dc:creator'), subject: g('dc:subject'), keywords: g('cp:keywords') };
  } catch { return {}; }
}
