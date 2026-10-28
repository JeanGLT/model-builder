/*
 * Minimal Markdown to HTML, enough for the term sheets: headings, paragraphs, tables,
 * bullet / numbered / checkbox lists, **bold** and *italic*. Text is HTML-escaped first.
 */
(function (root) {
  'use strict';

  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const inline = (s) => esc(s).replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>').replace(/\*(.+?)\*/g, '<em>$1</em>');
  const cells = (line) => line.trim().replace(/^\||\|$/g, '').split('|').map((c) => c.trim());

  function toHtml(md) {
    const lines = md.replace(/\r/g, '').split('\n');
    const out = [];
    let i = 0;
    while (i < lines.length) {
      const line = lines[i];
      if (!line.trim()) { i++; continue; }
      const h = line.match(/^(#{1,3})\s+(.*)$/);
      if (h) { out.push('<h' + h[1].length + '>' + inline(h[2]) + '</h' + h[1].length + '>'); i++; continue; }
      if (line.trim().startsWith('|')) {
        const head = cells(line);
        i += 2; // skip the | --- | separator row
        const rows = [];
        while (i < lines.length && lines[i].trim().startsWith('|')) rows.push(cells(lines[i++]));
        out.push('<div class="ts-table"><table><thead><tr>' + head.map((c) => '<th>' + inline(c) + '</th>').join('') + '</tr></thead><tbody>' +
          rows.map((r) => '<tr>' + r.map((c) => '<td>' + inline(c) + '</td>').join('') + '</tr>').join('') + '</tbody></table></div>');
        continue;
      }
      if (/^\s*(- |\d+\. )/.test(line)) {
        const ordered = /^\s*\d+\. /.test(line);
        const items = [];
        while (i < lines.length && /^\s*(- |\d+\. )/.test(lines[i])) {
          const t = lines[i++].replace(/^\s*(- |\d+\. )/, '');
          const box = t.match(/^\[( |x)\]\s*(.*)$/);
          items.push(box ? '<li class="todo"><span class="box">' + (box[1] === 'x' ? '✓' : '') + '</span>' + inline(box[2]) + '</li>' : '<li>' + inline(t) + '</li>');
        }
        out.push((ordered ? '<ol>' : '<ul>') + items.join('') + (ordered ? '</ol>' : '</ul>'));
        continue;
      }
      const para = [];
      while (i < lines.length && lines[i].trim() && !/^(#{1,3}\s|\s*\||\s*- |\s*\d+\. )/.test(lines[i])) para.push(lines[i++].trim());
      out.push('<p>' + inline(para.join(' ')) + '</p>');
    }
    return out.join('\n');
  }

  const api = { toHtml };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PFMarkdown = api;
})(typeof window !== 'undefined' ? window : globalThis);
