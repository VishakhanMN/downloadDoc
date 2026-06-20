import { Component, OnInit } from '@angular/core';
import { ViewDocument } from '../view-document/view-document';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  ShadingType,
  HeightRule,
  ExternalHyperlink,
  HighlightColor
} from 'docx';
import { saveAs } from 'file-saver';
import { FormsModule } from '@angular/forms';

@Component({
  selector: 'app-document',
  imports: [ViewDocument, FormsModule],
  templateUrl: './document.html',
  styleUrl: './document.css',
})
export class DocumentComponent implements OnInit {

  rollbackArtifactName: string = '';
  rollbackArtifactVersion: string = '';
  artifactName: string = '';
  artifactVersion: string = '';
  artifactBuildDate: string = '';
  buildJobPipelineId: string = '';
  gitBranch: string = '';
  tagNumber: string = '';
  deploymentVersion: string = '';
  jobId: string = '';
  requestedDate: string = '';

  /* 1 mthod to donwload without tables */
//   downloadDoc() {

//   const content = document.getElementById('content')?.innerHTML;

//   const header = `
//     <html xmlns:o='urn:schemas-microsoft-com:office:office'
//           xmlns:w='urn:schemas-microsoft-com:office:word'
//           xmlns='http://www.w3.org/TR/REC-html40'>
//       <head>
//         <meta charset='utf-8'>
//       </head>
//       <body>
//   `;

//   const footer = `
//       </body>
//     </html>
//   `;

//   const sourceHTML = header + content + footer;

//   const blob = new Blob(
//     ['\ufeff', sourceHTML],
//     { type: 'application/msword' }
//   );

//   const url = URL.createObjectURL(blob);

//   const link = document.createElement('a');

//   link.href = url;
//   link.download = 'document.doc';

//   document.body.appendChild(link);

//   link.click();

//   document.body.removeChild(link);
// }

ngOnInit() {
  const today = new Date();
  const dd = String(today.getDate()).padStart(2, '0');
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const yyyy = today.getFullYear();
  this.requestedDate = `${dd}-${mm}-${yyyy}`;
}

downloadDoc() {

  // Resolve CSS color class on an element, falling back to inherited color
  const getColor = (el: Element, inherited?: string): string | undefined => {
    if (el.classList.contains('text-red')) return 'FF0000';
    if (el.classList.contains('text-green')) return '008000';
    if (el.classList.contains('text-blue')) return '0D6EFD';
    return inherited;
  };

  // Resolve CSS background/highlight class on an element, falling back to inherited highlight
  type HighlightColorValue = (typeof HighlightColor)[keyof typeof HighlightColor];
  const getHighlight = (el: Element, inherited?: HighlightColorValue): HighlightColorValue | undefined => {
    if (el.classList.contains('background-yellow')) return HighlightColor.YELLOW;
    return inherited;
  };

  // Walk a DOM element and produce Paragraphs, splitting at <br>, preserving inline colors
  const nodeToParagraphs = (el: HTMLElement, inheritedColor: string | undefined, indentLeft?: number, inheritedHighlight?: HighlightColorValue, spacingBefore?: number): Paragraph[] => {
    const paragraphs: Paragraph[] = [];
    let currentRuns: (TextRun | ExternalHyperlink)[] = [];
    let isFirstParagraph = true;

    const flush = () => {
      if (currentRuns.length > 0) {
        paragraphs.push(new Paragraph({
          indent: indentLeft ? { left: indentLeft } : undefined,
          spacing: (spacingBefore && isFirstParagraph) ? { before: spacingBefore } : undefined,
          children: [...currentRuns],
        }));
        currentRuns = [];
        isFirstParagraph = false;
      }
    };

    const walk = (node: Node, color: string | undefined, highlight: HighlightColorValue | undefined) => {
      if (node.nodeType === Node.TEXT_NODE) {
        const text = (node.textContent ?? '').replace(/\s+/g, ' ').trim();
        if (text) currentRuns.push(new TextRun({ text, color, size: 22, ...(highlight ? { highlight } : {}) }));
        return;
      }
      if (node.nodeType !== Node.ELEMENT_NODE) return;
      const elem = node as HTMLElement;
      if (elem.tagName === 'BR') { flush(); return; }
      if (elem.tagName === 'UL') return; // handled by processListItems
      if (elem.tagName === 'A') {
        const href = elem.getAttribute('href')?.trim() ?? '';
        const text = elem.innerText.trim();
        if (href && text) {
          currentRuns.push(new ExternalHyperlink({
            link: href,
            children: [new TextRun({ text, style: 'Hyperlink', size: 22 })],
          }));
        }
        return;
      }
      if (elem.tagName === 'PRE') {
        flush();
        const preColor = getColor(elem, color);
        elem.innerText.split('\n').filter(l => l.trim()).forEach(line => {
          paragraphs.push(new Paragraph({
            indent: indentLeft ? { left: indentLeft } : undefined,
            children: [new TextRun({ text: line, color: preColor, size: 20 })],
          }));
        });
        return;
      }
      const elemColor = getColor(elem, color);
      const elemHighlight = getHighlight(elem, highlight);
      elem.childNodes.forEach(child => walk(child, elemColor, elemHighlight));
    };

    const rootColor = getColor(el, inheritedColor);
    const rootHighlight = getHighlight(el, inheritedHighlight);
    el.childNodes.forEach(child => walk(child, rootColor, rootHighlight));
    flush();
    return paragraphs.length > 0 ? paragraphs : [new Paragraph({ children: [new TextRun({ text: '', size: 22 })] })];
  };

  // Recursively convert <ul>/<li> to indented Paragraphs with color support
  const processListItems = (ul: HTMLElement, indent: number): Paragraph[] => {
    const result: Paragraph[] = [];
    Array.from(ul.querySelectorAll(':scope > li')).forEach(li => {
      const liEl = li as HTMLElement;
      const nestedUl = liEl.querySelector(':scope > ul');
      const spacingBefore = liEl.classList.contains('mt-4') ? 360 : undefined;
      result.push(...nodeToParagraphs(liEl, undefined, indent * 360, undefined, spacingBefore));
      if (nestedUl) {
        result.push(...processListItems(nestedUl as HTMLElement, indent + 1));
      }
    });
    return result;
  };

  // Build Paragraphs from a single table cell
  const cellToParagraphs = (cell: HTMLElement, isHeader: boolean, isFullWidth: boolean, isInThead: boolean): Paragraph[] => {
    if (isInThead) {
      const ps: Paragraph[] = [];
      const h2 = cell.querySelector('h2');
      const h4 = cell.querySelector('h4');
      if (h2) ps.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: (h2 as HTMLElement).innerText.trim(), bold: true, size: 36 })],
      }));
      if (h4) ps.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: (h4 as HTMLElement).innerText.trim(), italics: true, size: 22 })],
      }));
      return ps.length > 0 ? ps : [new Paragraph({ children: [new TextRun({ text: cell.innerText.trim(), size: 22 })] })];
    }

    // Cells containing a top-level <ul>
    const topUl = cell.querySelector(':scope > ul');
    if (topUl) {
      const ps = processListItems(topUl as HTMLElement, 0);
      return ps.length > 0 ? ps : [new Paragraph({ children: [new TextRun({ text: '', size: 22 })] })];
    }

    // Regular cells — DOM walk preserves <br> splits and inline colors
    return nodeToParagraphs(cell, undefined);
  };

  const tableEl = document.getElementById('myTable');

  if (!tableEl) {
    return;
  }

  const rows = Array.from(tableEl.querySelectorAll('tr'));

  const docxRows = rows
    .filter(tr => tr.querySelectorAll('th, td').length > 0)
    .map(tr => {
      const isInThead = tr.closest('thead') !== null;
      const cells = Array.from(tr.querySelectorAll('th, td'));
      const hasOnlyHeaders = cells.every(c => c.tagName.toLowerCase() === 'th');

      return new TableRow({
        height: { value: hasOnlyHeaders ? 600 : 500, rule: HeightRule.ATLEAST },
        children: cells.map(cell => {
          const isHeader = cell.tagName.toLowerCase() === 'th';
          const colspan = parseInt(cell.getAttribute('colspan') || '1', 10);
          const isFullWidth = colspan > 1;

          const paragraphs = cellToParagraphs(cell as HTMLElement, isHeader, isFullWidth, isInThead);

          return new TableCell({
            columnSpan: isFullWidth ? colspan : undefined,
            width: isFullWidth
              ? { size: 100, type: WidthType.PERCENTAGE }
              : { size: isHeader ? 15 : 85, type: WidthType.PERCENTAGE },
            shading: isHeader
              ? { fill: 'D3D3D3', type: ShadingType.CLEAR, color: 'auto' }
              : undefined,
            margins: { top: 120, bottom: 120, left: 180, right: 180 },
            children: paragraphs,
          });
        }),
      });
    });

  const docxTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: docxRows,
  });

  const doc = new Document({
    sections: [
      {
        children: [docxTable],
      },
    ],
  });

  Packer.toBlob(doc).then(blob => {
    saveAs(blob, `SIPS-SPA-STG_SPA-Deployment_${this.deploymentVersion}__Instructions.docx`);
  });
}

}
