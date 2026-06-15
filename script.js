// ==========================================
// CRK MARKDOWN CONVERTER
// Client-side MD <-> PDF / DOCX engine
// ==========================================

(function () {
    'use strict';

    // ------------------------------------------
    // LIBRARY READINESS GUARDS
    // ------------------------------------------
    const hasMarked = typeof marked !== 'undefined';
    const hasPurify = typeof DOMPurify !== 'undefined';
    const hasHljs = typeof hljs !== 'undefined';
    const hasMammoth = typeof mammoth !== 'undefined';
    const hasPdfjs = typeof pdfjsLib !== 'undefined';

    // Configure marked (modern marked ignores the old `highlight` option — we
    // highlight after render instead, so this works across marked versions).
    if (hasMarked && typeof marked.setOptions === 'function') {
        marked.setOptions({ breaks: true, gfm: true });
    }

    // Configure Turndown (HTML -> Markdown) for imports. The GFM plugin is what
    // makes tables, strikethrough and task lists convert accurately — plain
    // Turndown silently drops tables.
    let turndownService = null;
    if (typeof TurndownService !== 'undefined') {
        turndownService = new TurndownService({
            headingStyle: 'atx',
            codeBlockStyle: 'fenced',
            bulletListMarker: '-',
            emDelimiter: '*',
            strongDelimiter: '**',
            linkStyle: 'inlined'
        });

        if (typeof turndownPluginGfm !== 'undefined') {
            turndownService.use(turndownPluginGfm.gfm);
        }

        // Preserve hard line breaks (common in Word documents).
        turndownService.addRule('hardBreak', {
            filter: 'br',
            replacement: function () { return '  \n'; }
        });

        // Drop empty paragraphs Word likes to insert so output stays clean.
        turndownService.addRule('dropEmptyParagraphs', {
            filter: function (node) {
                return node.nodeName === 'P' && node.textContent.trim() === '' && !node.querySelector('img');
            },
            replacement: function () { return ''; }
        });
    }

    // PDF.js worker is initialised lazily as a SAME-ORIGIN Blob (see
    // ensurePdfWorker) to avoid the cross-origin Worker security error that
    // occurs when pointing workerSrc straight at a CDN URL.
    const PDF_WORKER_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

    // ------------------------------------------
    // DOM ELEMENTS
    // ------------------------------------------
    const $ = (id) => document.getElementById(id);

    const editor = $('editor');
    const preview = $('preview');
    const fileInput = $('fileInput');
    const dropZone = $('dropZone');
    const wordCountEl = $('wordCount');
    const charCountEl = $('charCount');
    const readTimeEl = $('readTime');
    const loadingOverlay = $('loadingOverlay');
    const loadingText = $('loadingText');
    const toastContainer = $('toastContainer');

    const importBtn = $('importBtn');
    const exportPdfBtn = $('exportPdfBtn');
    const exportDocxBtn = $('exportDocxBtn');
    const exportMdBtn = $('exportMdBtn');
    const copyMarkdownBtn = $('copyMarkdownBtn');
    const clearBtn = $('clearBtn');

    let currentFileName = 'crk-export';

    const DEFAULT_CONTENT = `# Welcome to CRK Converter ⚡

This is a Markdown to PDF/DOCX converter running purely in your browser.

## Features

- [x] Live Preview
- [x] Drag & Drop files
- [x] Export to PDF, DOCX & MD
- [x] Syntax Highlighting

\`\`\`javascript
// Pure client-side magic
function engageHyperdrive() {
  console.log("System optimal.");
}
\`\`\`

> "In a world of users,be the one who programs."

| Format | Support | Status |
|--------|---------|--------|
| MD     | Full    | Online |
| PDF    | Export  | Online |
| DOCX   | Export  | Online |
`;

    // ------------------------------------------
    // UI HELPERS
    // ------------------------------------------
    function showToast(message, type) {
        if (!toastContainer) { console.log(message); return; }
        const toast = document.createElement('div');
        toast.className = 'toast toast-' + (type || 'info');
        toast.textContent = message;
        toastContainer.appendChild(toast);
        // Force reflow then animate in.
        requestAnimationFrame(() => toast.classList.add('show'));
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.remove(), 300);
        }, 3200);
    }

    function showLoading(text) {
        if (loadingText && text) loadingText.textContent = text;
        if (loadingOverlay) loadingOverlay.style.display = 'flex';
    }

    function hideLoading() {
        if (loadingOverlay) loadingOverlay.style.display = 'none';
    }

    function escapeHtml(str) {
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
    }

    function downloadBlob(blob, filename) {
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        // Revoke a little later so the download has time to start.
        setTimeout(() => URL.revokeObjectURL(url), 1500);
    }

    function isMobileDevice() {
        return /Android|iPhone|iPad|iPod|Mobi/i.test(navigator.userAgent) ||
            // iPadOS 13+ reports as desktop Safari, so detect it via touch points.
            (navigator.maxTouchPoints > 1 && /Mac/.test(navigator.platform));
    }

    // Save a blob. On phones/tablets the classic <a download> is unreliable —
    // iOS Safari ignores it for blob URLs — so we use the Web Share API, which
    // pops the native sheet to save into Files / Word / Drive / Mail. Desktop
    // keeps the silent direct download. (Web Share needs a secure context:
    // works on HTTPS and on localhost, NOT on a plain http LAN IP.)
    async function saveFile(blob, filename, mime) {
        if (isMobileDevice() && typeof navigator.canShare === 'function') {
            try {
                const file = new File([blob], filename, { type: mime || blob.type || 'application/octet-stream' });
                if (navigator.canShare({ files: [file] })) {
                    await navigator.share({ files: [file], title: filename });
                    return 'shared';
                }
            } catch (err) {
                if (err && err.name === 'AbortError') return 'cancelled';
                console.warn('Web Share failed; falling back to direct download.', err);
            }
        }
        downloadBlob(blob, filename);
        return 'downloaded';
    }

    // ------------------------------------------
    // CORE: RENDER & STATS
    // ------------------------------------------
    function renderMarkdown(markdownText) {
        let html;
        try {
            if (hasMarked && typeof marked.parse === 'function') {
                html = marked.parse(markdownText || '');
            } else if (hasMarked && typeof marked === 'function') {
                html = marked(markdownText || '');
            } else {
                html = '<pre>' + escapeHtml(markdownText || '') + '</pre>';
            }
        } catch (e) {
            console.error('Markdown render failed:', e);
            html = '<pre>' + escapeHtml(markdownText || '') + '</pre>';
        }

        if (hasPurify) {
            try { html = DOMPurify.sanitize(html, { ADD_ATTR: ['target'] }); } catch (e) { /* keep raw */ }
        }

        preview.innerHTML = html;

        // Apply syntax highlighting on freshly rendered code blocks.
        if (hasHljs) {
            preview.querySelectorAll('pre code').forEach((block) => {
                try { hljs.highlightElement(block); } catch (e) { /* ignore */ }
            });
        }
    }

    function updateStats(text) {
        const chars = text.length;
        if (charCountEl) charCountEl.textContent = chars.toLocaleString();

        const words = text.trim() === '' ? 0 : text.trim().split(/\s+/).length;
        if (wordCountEl) wordCountEl.textContent = words.toLocaleString();

        const readTime = Math.max(words === 0 ? 0 : 1, Math.ceil(words / 200));
        if (readTimeEl) readTimeEl.textContent = readTime;
    }

    // Stats + persistence update instantly; the heavier markdown render is
    // debounced so large documents stay smooth while typing.
    let renderTimer = null;
    function onEditorInput() {
        const text = editor.value;
        updateStats(text);
        try { localStorage.setItem('crk_content', text); } catch (e) { /* quota */ }
        clearTimeout(renderTimer);
        renderTimer = setTimeout(() => renderMarkdown(text), 120);
    }

    function setEditorContent(content) {
        editor.value = content || '';
        updateStats(editor.value);
        try { localStorage.setItem('crk_content', editor.value); } catch (e) { /* quota */ }
        renderMarkdown(editor.value);
    }

    // ------------------------------------------
    // IMPORT ENGINE (MD / TXT / PDF / DOCX -> MD)
    // ------------------------------------------
    async function handleFile(file) {
        if (!file) return;
        const extension = (file.name.split('.').pop() || '').toLowerCase();
        const base = file.name.substring(0, file.name.lastIndexOf('.'));
        currentFileName = base || 'crk-export';

        showLoading('IMPORTING FILE...');
        try {
            if (extension === 'md' || extension === 'markdown' || extension === 'txt') {
                const text = await file.text();
                setEditorContent(text);
                showToast('Imported ' + file.name, 'success');
            } else if (extension === 'docx' || extension === 'doc') {
                if (!turndownService) throw new Error('Turndown library unavailable.');
                const arrayBuffer = await file.arrayBuffer();
                let html = null;

                if (hasMammoth) {
                    try {
                        const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
                        html = result.value;
                    } catch (mErr) {
                        // Some ".doc/.docx" files are actually saved HTML, not real
                        // OOXML — fall through to the HTML reader below.
                        console.warn('mammoth could not parse as OOXML, trying HTML fallback.', mErr);
                    }
                }

                if (!html || html.trim() === '') {
                    const asText = await file.text();
                    if (/<\/?(html|body|table|p|h[1-6]|div|span|ul|ol|li)\b/i.test(asText)) {
                        html = asText;
                    } else {
                        throw new Error('Not a readable Word document.');
                    }
                }

                const markdown = turndownService.turndown(html);
                setEditorContent(markdown);
                showToast('Imported ' + file.name, 'success');
            } else if (extension === 'pdf') {
                if (!hasPdfjs) throw new Error('PDF import library unavailable.');
                const arrayBuffer = await file.arrayBuffer();
                const text = await extractTextFromPDF(arrayBuffer);
                setEditorContent('# ' + (base || 'Extracted PDF Content') + '\n\n' + text);
                showToast('Imported ' + file.name, 'success');
            } else {
                showToast('Unsupported format. Use .md, .txt, .docx or .pdf', 'error');
            }
        } catch (error) {
            console.error('Error processing file:', error);
            const msg = (error && error.message) ? error.message : 'unknown error';
            showToast('Import failed: ' + msg, 'error');
        } finally {
            hideLoading();
        }
    }

    // Load the PDF.js worker as a same-origin Blob (cdnjs serves it with CORS),
    // which sidesteps the "Worker scripts must be same-origin" error.
    let pdfWorkerPromise = null;
    function ensurePdfWorker() {
        if (pdfWorkerPromise) return pdfWorkerPromise;
        pdfWorkerPromise = fetch(PDF_WORKER_URL)
            .then((r) => { if (!r.ok) throw new Error('worker HTTP ' + r.status); return r.text(); })
            .then((code) => {
                const blob = new Blob([code], { type: 'application/javascript' });
                pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(blob);
            })
            .catch((err) => {
                console.warn('Blob worker load failed; falling back to direct URL.', err);
                pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_URL;
            });
        return pdfWorkerPromise;
    }

    // Build a GFM markdown table from an array of rows (each row = array of cell
    // strings). Column count is normalised to the widest row.
    function renderMarkdownTable(rows) {
        const colCount = rows.reduce((m, r) => Math.max(m, r.length), 0);
        if (colCount < 2) return rows.map((r) => r.join(' ')).join('\n');

        const esc = (t) => (t || '').replace(/\|/g, '\\|').replace(/\s*\n\s*/g, ' ').trim();
        const toRow = (cells) => {
            const out = [];
            for (let c = 0; c < colCount; c++) out.push(esc(cells[c] || ''));
            return '| ' + out.join(' | ') + ' |';
        };

        const header = toRow(rows[0]);
        const sep = '| ' + new Array(colCount).fill('---').join(' | ') + ' |';
        const body = rows.slice(1).map(toRow);
        return [header, sep].concat(body).join('\n');
    }

    // Extract text from a PDF, reconstructing paragraphs AND tables from the
    // x/y geometry of each glyph run (PDFs have no semantic structure, so this
    // is inferred). Emojis and Unicode pass through as-is.
    async function extractTextFromPDF(arrayBuffer) {
        await ensurePdfWorker();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        const pageBlocks = [];

        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();

            // 1) Group glyph runs into visual rows by Y position.
            const rows = [];
            let current = null;
            let lastY = null;
            for (const item of textContent.items) {
                if (!item.str) continue;
                const x = item.transform[4];
                const y = item.transform[5];
                const h = item.height || Math.abs(item.transform[3]) || 10;
                const seg = { x: x, end: x + (item.width || 0), str: item.str, h: h };
                if (lastY === null || Math.abs(y - lastY) < 3) {
                    if (!current) current = { y: y, gapBefore: 0, segs: [] };
                    current.segs.push(seg);
                } else {
                    rows.push(current);
                    current = { y: y, gapBefore: lastY - y, segs: [seg] };
                }
                lastY = y;
            }
            if (current) rows.push(current);

            // 2) Split each row into cells using significant horizontal gaps
            //    (relative to font size, so word spaces are not mistaken for cells).
            const processed = rows.map((row) => {
                row.segs.sort((a, b) => a.x - b.x);
                const cells = [];
                let cellText = '';
                let prevEnd = null;
                for (const seg of row.segs) {
                    const fontGap = Math.max(10, seg.h * 1.4);
                    if (prevEnd !== null && (seg.x - prevEnd) > fontGap) {
                        cells.push(cellText.trim());
                        cellText = '';
                    } else if (cellText && prevEnd !== null && (seg.x - prevEnd) > seg.h * 0.2 &&
                               !cellText.endsWith(' ') && !/^\s/.test(seg.str)) {
                        cellText += ' ';
                    }
                    cellText += seg.str;
                    prevEnd = seg.end;
                }
                if (cellText.trim() !== '') cells.push(cellText.trim());
                return { gapBefore: row.gapBefore, cells: cells };
            }).filter((r) => r.cells.length > 0);

            // 3) Group consecutive multi-column rows into tables; everything else
            //    becomes paragraph text (blank line on large vertical gaps).
            const out = [];
            let tableBuf = [];
            const flush = () => {
                if (!tableBuf.length) return;
                if (tableBuf.length >= 2) {
                    out.push('');
                    out.push(renderMarkdownTable(tableBuf.map((r) => r.cells)));
                    out.push('');
                } else {
                    out.push(tableBuf[0].cells.join(' '));
                }
                tableBuf = [];
            };
            for (const r of processed) {
                if (r.cells.length >= 2) {
                    tableBuf.push(r);
                } else {
                    flush();
                    if (r.gapBefore > 14 && out.length) out.push('');
                    out.push(r.cells[0]);
                }
            }
            flush();

            pageBlocks.push(out.join('\n').replace(/\n{3,}/g, '\n\n').trim());
        }

        return pageBlocks.join('\n\n---\n\n');
    }

    // ------------------------------------------
    // SHARED: clean light-themed document CSS used by PDF (print) and DOCX.
    // ------------------------------------------
    function exportStyles(forPrint) {
        return `
        ${forPrint ? '@page { size: A4; margin: 16mm; }' : ''}
        body { font-family: 'Calibri','Arial',sans-serif; font-size: ${forPrint ? '12pt' : '11pt'}; line-height: 1.55; color: #1a1a1a; background: #ffffff; ${forPrint ? 'margin: 0;' : ''} }
        h1, h2, h3, h4, h5, h6 { font-family: 'Calibri','Arial',sans-serif; color: #000000; margin-top: 14pt; margin-bottom: 7pt; page-break-after: avoid; }
        h1 { font-size: 20pt; border-bottom: 1px solid #cccccc; padding-bottom: 3pt; }
        h2 { font-size: 16pt; }
        h3 { font-size: 13pt; }
        p, li { margin-bottom: 6pt; }
        ul, ol { padding-left: 26px; }
        a { color: #0055cc; text-decoration: underline; }
        img { max-width: 100%; }
        pre { background: #f6f8fa; border: 1px solid #dddddd; border-radius: 4px; padding: 10pt; font-family: 'Consolas','Courier New',monospace; font-size: 10pt; white-space: pre-wrap; word-break: break-word; page-break-inside: avoid; }
        code { font-family: 'Consolas','Courier New',monospace; font-size: 10pt; background: #f6f8fa; color: #d73a49; padding: 2px 4px; border-radius: 3px; }
        pre code { color: #24292e; background: transparent; padding: 0; }
        blockquote { border-left: 4px solid #dddddd; background: #f9f9f9; color: #555555; padding: 8pt 12pt; margin: 0 0 12pt 0; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 12pt; page-break-inside: avoid; }
        th, td { border: 1px solid #bbbbbb; padding: 6pt 9pt; text-align: left; }
        th { background: #f2f2f2; font-weight: bold; color: #000000; }
        /* syntax highlighting (light palette) */
        .hljs-keyword,.hljs-selector-tag,.hljs-subst { color: #d73a49; font-weight: bold; }
        .hljs-string,.hljs-doctag,.hljs-template-variable,.hljs-addition { color: #032f62; }
        .hljs-comment,.hljs-quote { color: #6a737d; font-style: italic; }
        .hljs-number,.hljs-literal,.hljs-variable,.hljs-template-tag { color: #005cc5; }
        .hljs-title,.hljs-section,.hljs-selector-id { color: #6f42c1; }
        .hljs-type,.hljs-built_in { color: #e36209; }
        `;
    }

    // ------------------------------------------
    // EXPORT: PDF (via the browser's native print engine)
    // Opens a clean, light-themed print window and triggers print -> "Save as
    // PDF". This produces real selectable text with automatic multi-page
    // pagination, and never yields the blank/one-page output that screenshot-
    // based engines (html2canvas) produce on some browsers.
    // ------------------------------------------
    function exportToPDF() {
        if (!editor.value.trim()) { showToast('Nothing to export yet.', 'error'); return; }

        const docHtml = `<!DOCTYPE html><html><head><meta charset="utf-8">
<title>${currentFileName}</title>
<style>${exportStyles(true)}</style>
</head><body class="markdown-body">${preview.innerHTML}
<script>window.onload=function(){setTimeout(function(){window.focus();window.print();},250);};<\/script>
</body></html>`;

        const win = window.open('', '_blank');
        if (!win) {
            showToast('Allow pop-ups for this site, then try again.', 'error');
            return;
        }
        win.document.open();
        win.document.write(docHtml);
        win.document.close();
        showToast("In the dialog, choose 'Save as PDF'.", 'info');
    }

    // ------------------------------------------
    // EXPORT: WORD (.doc)
    // We deliberately emit a Word-compatible HTML document (.doc) rather than a
    // packaged .docx. html-docx-js embeds the body as an OOXML "altChunk", which
    // ONLY desktop MS Word knows how to expand — Google Docs and the Word mobile
    // apps render that as a BLANK page. A Word-HTML .doc opens with full content
    // everywhere: desktop Word, Word mobile, Google Docs, LibreOffice.
    // ------------------------------------------
    function buildExportHtml() {
        let body = preview.innerHTML;

        // Add Word-friendly table styling.
        body = body
            .replace(/<table/g, '<table border="1" cellspacing="0" cellpadding="6" style="border-collapse: collapse; width: 100%; border: 1px solid #000000;"')
            .replace(/<th/g, '<th style="border: 1px solid #000000; padding: 6px; background-color: #f2f2f2; font-weight: bold;"')
            .replace(/<td/g, '<td style="border: 1px solid #000000; padding: 6px;"');

        // The mso block gives Word real A4 page geometry/margins on open.
        const head = `<!DOCTYPE html><html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head>
    <meta charset='utf-8'>
    <title>${escapeHtml(currentFileName)}</title>
    <!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom></w:WordDocument></xml><![endif]-->
    <style>
        @page { size: A4; margin: 2cm; }
        ${exportStyles(false)}
    </style>
</head>
<body>`;
        return head + body + '</body></html>';
    }

    // ------------------------------------------
    // REAL .docx (OOXML) GENERATION
    // Builds a genuine Office Open XML package with JSZip so the file opens
    // natively in Google Docs and the Word apps on Android/iOS. (The old
    // altChunk / Word-HTML approaches render blank in non-desktop viewers.)
    // ------------------------------------------
    function dxEsc(s) {
        return String(s)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;')
            .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    // One run of text with optional formatting; newlines become <w:br/>.
    function dxRun(text, fmt) {
        fmt = fmt || {};
        if (text === '' || text == null) return '';
        let rpr = '';
        if (fmt.bold) rpr += '<w:b/>';
        if (fmt.italic) rpr += '<w:i/>';
        if (fmt.strike) rpr += '<w:strike/>';
        if (fmt.code) rpr += '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:color w:val="C7254E"/>';
        if (fmt.link) rpr += '<w:color w:val="0563C1"/><w:u w:val="single"/>';
        if (fmt.color && !fmt.code && !fmt.link) rpr += '<w:color w:val="' + fmt.color + '"/>';
        if (fmt.sz) rpr += '<w:sz w:val="' + fmt.sz + '"/><w:szCs w:val="' + fmt.sz + '"/>';
        const rprXml = rpr ? '<w:rPr>' + rpr + '</w:rPr>' : '';
        const parts = String(text).split('\n');
        let inner = '';
        for (let i = 0; i < parts.length; i++) {
            if (i > 0) inner += '<w:br/>';
            inner += '<w:t xml:space="preserve">' + dxEsc(parts[i]) + '</w:t>';
        }
        return '<w:r>' + rprXml + inner + '</w:r>';
    }

    // Walk inline child nodes accumulating formatted runs.
    function dxInline(node, fmt) {
        fmt = fmt || {};
        let out = '';
        node.childNodes.forEach(function (child) {
            if (child.nodeType === 3) { out += dxRun(child.nodeValue, fmt); return; }
            if (child.nodeType !== 1) return;
            const tag = child.tagName.toLowerCase();
            if (tag === 'br') { out += '<w:r><w:br/></w:r>'; return; }
            if (tag === 'input') { out += dxRun(child.checked ? '☑ ' : '☐ ', fmt); return; }
            if (tag === 'img') { out += dxRun('[image: ' + (child.getAttribute('alt') || '') + ']', fmt); return; }
            const f = Object.assign({}, fmt);
            if (tag === 'strong' || tag === 'b') f.bold = true;
            else if (tag === 'em' || tag === 'i') f.italic = true;
            else if (tag === 'del' || tag === 's' || tag === 'strike') f.strike = true;
            else if (tag === 'code') f.code = true;
            else if (tag === 'a') f.link = true;
            out += dxInline(child, f);
        });
        return out;
    }

    function dxPara(runsXml, pprInner) {
        const ppr = pprInner ? '<w:pPr>' + pprInner + '</w:pPr>' : '';
        return '<w:p>' + ppr + (runsXml || '') + '</w:p>';
    }

    // Lists become indented paragraphs with bullet/number prefixes (no
    // numbering.xml needed — renders correctly in every viewer).
    function dxList(listNode, ordered, depth) {
        let xml = '';
        let idx = 1;
        listNode.childNodes.forEach(function (li) {
            if (li.nodeType !== 1 || li.tagName.toLowerCase() !== 'li') return;
            const clone = li.cloneNode(true);
            Array.prototype.slice.call(clone.querySelectorAll('ul,ol')).forEach(function (n) {
                n.parentNode.removeChild(n);
            });
            const prefix = ordered ? (idx + '. ') : '• ';
            idx++;
            const ind = 360 + depth * 360;
            xml += dxPara(dxRun(prefix, {}) + dxInline(clone, {}), '<w:ind w:left="' + ind + '"/>');
            li.childNodes.forEach(function (c) {
                if (c.nodeType === 1 && /^(ul|ol)$/.test(c.tagName.toLowerCase())) {
                    xml += dxList(c, c.tagName.toLowerCase() === 'ol', depth + 1);
                }
            });
        });
        return xml;
    }

    function dxTable(table) {
        const sides = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
        let borders = '<w:tblBorders>';
        sides.forEach(function (b) { borders += '<w:' + b + ' w:val="single" w:sz="4" w:space="0" w:color="999999"/>'; });
        borders += '</w:tblBorders>';
        let rowsXml = '';
        Array.prototype.slice.call(table.querySelectorAll('tr')).forEach(function (tr) {
            let cellsXml = '';
            tr.childNodes.forEach(function (cell) {
                if (cell.nodeType !== 1) return;
                const tg = cell.tagName.toLowerCase();
                if (tg !== 'td' && tg !== 'th') return;
                const isHead = tg === 'th';
                const shd = isHead ? '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>' : '';
                cellsXml += '<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/>' + shd + '</w:tcPr>' +
                    dxPara(dxInline(cell, { bold: isHead }), '') + '</w:tc>';
            });
            if (cellsXml) rowsXml += '<w:tr>' + cellsXml + '</w:tr>';
        });
        return '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>' + borders + '</w:tblPr>' + rowsXml + '</w:tbl>' + dxPara('', '');
    }

    function dxBlocks(container) {
        const HSIZE = { h1: 36, h2: 32, h3: 28, h4: 26, h5: 24, h6: 22 };
        let xml = '';
        container.childNodes.forEach(function (node) {
            if (node.nodeType === 3) {
                if (node.nodeValue.trim() !== '') xml += dxPara(dxRun(node.nodeValue, {}), '');
                return;
            }
            if (node.nodeType !== 1) return;
            const tag = node.tagName.toLowerCase();
            if (HSIZE[tag]) {
                xml += dxPara(dxInline(node, { bold: true, sz: HSIZE[tag] }), '<w:spacing w:before="240" w:after="120"/>');
            } else if (tag === 'p') {
                xml += dxPara(dxInline(node, {}), '<w:spacing w:after="160"/>');
            } else if (tag === 'ul' || tag === 'ol') {
                xml += dxList(node, tag === 'ol', 0);
            } else if (tag === 'blockquote') {
                node.childNodes.forEach(function (c) {
                    if (c.nodeType === 1) {
                        xml += dxPara(dxInline(c, { italic: true, color: '555555' }),
                            '<w:ind w:left="480"/><w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="CCCCCC"/></w:pBdr>');
                    } else if (c.nodeType === 3 && c.nodeValue.trim() !== '') {
                        xml += dxPara(dxRun(c.nodeValue, { italic: true, color: '555555' }), '<w:ind w:left="480"/>');
                    }
                });
            } else if (tag === 'pre') {
                const codeText = node.textContent.replace(/\n+$/, '');
                xml += dxPara(dxRun(codeText, { code: true, sz: 20 }), '<w:shd w:val="clear" w:color="auto" w:fill="F6F8FA"/>');
            } else if (tag === 'table') {
                xml += dxTable(node);
            } else if (tag === 'hr') {
                xml += dxPara('', '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="CCCCCC"/></w:pBdr>');
            } else {
                xml += dxBlocks(node);
            }
        });
        return xml;
    }

    // Assemble the OOXML package and return a Promise<Blob>.
    function buildDocxBlob() {
        let body = dxBlocks(preview);
        if (!body) body = '<w:p/>';
        const sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
        const documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
            body + sect + '</w:body></w:document>';

        const contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
            '<Default Extension="xml" ContentType="application/xml"/>' +
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
            '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
            '</Types>';

        const rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
            '</Relationships>';

        const docRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
            '</Relationships>';

        const stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
            '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>' +
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
            '</w:styles>';

        const zip = new JSZip();
        zip.file('[Content_Types].xml', contentTypes);
        zip.folder('_rels').file('.rels', rels);
        const wf = zip.folder('word');
        wf.file('document.xml', documentXml);
        wf.file('styles.xml', stylesXml);
        wf.folder('_rels').file('document.xml.rels', docRels);
        return zip.generateAsync({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            compression: 'DEFLATE'
        });
    }

    async function exportToDOCX() {
        if (!editor.value.trim()) { showToast('Nothing to export yet.', 'error'); return; }

        try {
            // Preferred: a REAL .docx (OOXML) — opens natively in Google Docs and
            // the Word apps on Android/iOS, where the Word-HTML .doc shows blank.
            if (typeof JSZip !== 'undefined') {
                const docxBlob = await buildDocxBlob();
                const howX = await saveFile(docxBlob, currentFileName + '.docx',
                    'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
                if (howX === 'cancelled') return;
                showToast(howX === 'shared' ? 'Choose where to save the .docx file.' : 'Word (.docx) exported.', 'success');
                return;
            }

            // Fallback (JSZip unavailable): Word-compatible HTML saved as .doc.
            const sourceHTML = buildExportHtml();
            // Leading U+FEFF BOM + application/msword makes every Word reader treat
            // the HTML as a document and decode UTF-8 (so emojis survive).
            const blob = new Blob(['﻿', sourceHTML], { type: 'application/msword' });
            const how = await saveFile(blob, currentFileName + '.doc', 'application/msword');
            if (how === 'cancelled') return;
            showToast(how === 'shared' ? 'Choose where to save the Word file.' : 'Word document exported.', 'success');
        } catch (err) {
            console.error('Word generation failed:', err);
            showToast('Word export failed. See console.', 'error');
        }
    }

    // ------------------------------------------
    // EXPORT: MARKDOWN
    // ------------------------------------------
    async function exportToMarkdown() {
        if (!editor.value.trim()) { showToast('Nothing to export yet.', 'error'); return; }
        const blob = new Blob([editor.value], { type: 'text/markdown;charset=utf-8' });
        const how = await saveFile(blob, currentFileName + '.md', 'text/markdown');
        if (how === 'cancelled') return;
        showToast(how === 'shared' ? 'Choose where to save the Markdown file.' : 'Markdown saved.', 'success');
    }

    // ------------------------------------------
    // EVENT WIRING
    // ------------------------------------------
    if (editor) editor.addEventListener('input', onEditorInput);

    if (importBtn && fileInput) {
        importBtn.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) handleFile(e.target.files[0]);
            fileInput.value = '';
        });
    }

    if (exportPdfBtn) exportPdfBtn.addEventListener('click', exportToPDF);
    if (exportDocxBtn) exportDocxBtn.addEventListener('click', exportToDOCX);
    if (exportMdBtn) exportMdBtn.addEventListener('click', exportToMarkdown);

    if (copyMarkdownBtn) {
        copyMarkdownBtn.addEventListener('click', () => {
            navigator.clipboard.writeText(editor.value).then(() => {
                const original = copyMarkdownBtn.textContent;
                copyMarkdownBtn.textContent = 'Copied!';
                setTimeout(() => { copyMarkdownBtn.textContent = original; }, 2000);
            }).catch(() => showToast('Clipboard access denied.', 'error'));
        });
    }

    if (clearBtn) {
        clearBtn.addEventListener('click', () => {
            if (confirm('Clear the editor? This action cannot be undone.')) {
                setEditorContent('');
            }
        });
    }

    // Drag & Drop (counter avoids flicker when moving over child elements).
    if (dropZone) {
        let dragDepth = 0;
        dropZone.addEventListener('dragenter', (e) => {
            e.preventDefault();
            dragDepth++;
            dropZone.classList.add('dragover');
        });
        dropZone.addEventListener('dragover', (e) => e.preventDefault());
        dropZone.addEventListener('dragleave', (e) => {
            e.preventDefault();
            dragDepth = Math.max(0, dragDepth - 1);
            if (dragDepth === 0) dropZone.classList.remove('dragover');
        });
        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dragDepth = 0;
            dropZone.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
        });
    }

    // Prevent the browser from navigating away (opening the file) when a file is
    // dropped anywhere outside the drop zone — route it to the importer instead.
    window.addEventListener('dragover', (e) => e.preventDefault());
    window.addEventListener('drop', (e) => {
        if (dropZone && dropZone.contains(e.target)) return; // handled above
        e.preventDefault();
        if (e.dataTransfer && e.dataTransfer.files.length > 0) {
            handleFile(e.dataTransfer.files[0]);
        }
    });

    // Synchronized scrolling between editor and preview.
    if (editor && preview && preview.parentElement) {
        const previewPane = preview.parentElement;
        let lock = false;

        const sync = (source, target) => {
            if (lock) return;
            lock = true;
            const max = source.scrollHeight - source.clientHeight;
            const ratio = max > 0 ? source.scrollTop / max : 0;
            target.scrollTop = ratio * (target.scrollHeight - target.clientHeight);
            requestAnimationFrame(() => { lock = false; });
        };

        editor.addEventListener('scroll', () => sync(editor, previewPane));
        previewPane.addEventListener('scroll', () => sync(previewPane, editor));
    }

    // ------------------------------------------
    // INITIAL LOAD
    // ------------------------------------------
    function init() {
        let saved = null;
        try { saved = localStorage.getItem('crk_content'); } catch (e) { /* ignore */ }
        editor.value = (saved !== null && saved !== '') ? saved : DEFAULT_CONTENT;

        updateStats(editor.value);
        renderMarkdown(editor.value);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
