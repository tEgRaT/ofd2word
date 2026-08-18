const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

/**
 * Pure Node.js Multi-Page OFD to DOCX Converter
 * Compatible with Node.js v13.6.0+ and legacy Windows
 */
class OFD2WordConverter {
    constructor(ofdPath, outputDocxPath) {
        this.ofdPath = path.resolve(ofdPath);
        this.outputDocxPath = path.resolve(outputDocxPath);
        this.tempDir = path.join(process.cwd(), `ofd_tmp_${Date.now()}`);
        this.docxDir = path.join(process.cwd(), `docx_tmp_${Date.now()}`);
        this.imagesToEmbed = [];
        this.imgIdx = 1; // Global image counter across all pages
    }

    // Unzip OFD archive
    unzipOFD() {
        fs.mkdirSync(this.tempDir, { recursive: true });

        if (process.platform === 'win32') {
            const tempZipPath = path.join(process.cwd(), `temp_in_${Date.now()}.zip`);
            fs.copyFileSync(this.ofdPath, tempZipPath);

            const winZip = path.resolve(tempZipPath).replace(/\//g, '\\');
            const winDest = path.resolve(this.tempDir).replace(/\//g, '\\');

            try {
                const psScript = `$s = New-Object -ComObject Shell.Application; $z = $s.NameSpace('${winZip}'); $d = $s.NameSpace('${winDest}'); if(-not $z -or -not $d){ exit 1 }; $items = $z.Items(); $d.CopyHere($items, 16); $start = Get-Date; while(($d.Items().Count -lt $items.Count) -and ((Get-Date) - $start).TotalSeconds -lt 10){ Start-Sleep -m 100 }`;
                execSync(`powershell -NoProfile -NonInteractive -Command "${psScript}"`);
            } finally {
                if (fs.existsSync(tempZipPath)) {
                    fs.unlinkSync(tempZipPath);
                }
            }
        } else {
            execSync(`unzip -q "${this.ofdPath}" -d "${this.tempDir}"`);
        }
    }

    // Discover all page Content.xml files in correct document order
    getPagePaths() {
        const pagesDir = path.join(this.tempDir, 'Doc_0', 'Pages');
        if (!fs.existsSync(pagesDir)) return [];

        const pagePaths = [];
        const docXmlPath = path.join(this.tempDir, 'Doc_0', 'Document.xml');

        // Method 1: Parse Document.xml for canonical page manifest
        if (fs.existsSync(docXmlPath)) {
            const xmlContent = fs.readFileSync(docXmlPath, 'utf-8');
            const pageNodes = xmlContent.match(/<(?:[a-zA-Z0-9]+:)?Page\b[\s\S]*?(?:\/>|><\/(?:[a-zA-Z0-9]+:)?Page>)/g) || [];

            for (const node of pageNodes) {
                const locMatch = node.match(/\bBaseLoc="([^"]+)"/) || node.match(/\bBaseLoc='([^']+)'/);
                if (locMatch) {
                    const fullPath = path.join(this.tempDir, 'Doc_0', locMatch[1]);
                    if (fs.existsSync(fullPath)) {
                        pagePaths.push(fullPath);
                    }
                }
            }
        }

        // Method 2: Fallback filesystem directory scan (sorted numerically)
        if (pagePaths.length === 0) {
            const entries = fs.readdirSync(pagesDir, { withFileTypes: true });
            const pageDirs = entries
                .filter(e => e.isDirectory() && /^Page_\d+$/i.test(e.name))
                .sort((a, b) => {
                    const numA = parseInt(a.name.replace(/Page_/i, ''), 10) || 0;
                    const numB = parseInt(b.name.replace(/Page_/i, ''), 10) || 0;
                    return numA - numB;
                });

            for (const dir of pageDirs) {
                const contentPath = path.join(pagesDir, dir.name, 'Content.xml');
                if (fs.existsSync(contentPath)) {
                    pagePaths.push(contentPath);
                }
            }
        }

        return pagePaths;
    }

    // Extract Font Names and Font-level styles
    extractFontMap() {
        const fontMap = {};
        const resFiles = [
            path.join(this.tempDir, 'Doc_0', 'PublicRes.xml'),
            path.join(this.tempDir, 'Doc_0', 'DocumentRes.xml')
        ];

        for (const resPath of resFiles) {
            if (!fs.existsSync(resPath)) continue;

            const xmlContent = fs.readFileSync(resPath, 'utf-8');
            const fontNodes = xmlContent.match(/<(?:[a-zA-Z0-9]+:)?Font\b[\s\S]*?>/g) || [];

            for (const node of fontNodes) {
                const idMatch = node.match(/\bID="([^"]+)"/) || node.match(/\bID='([^']+)'/);
                const nameMatch = node.match(/\bFontName="([^"]+)"/) || node.match(/\bFamilyName="([^"]+)"/) ||
                    node.match(/\bFontName='([^']+)'/) || node.match(/\bFamilyName='([^']+)'/);
                const weightMatch = node.match(/\bWeight=["']([^"']+)["']/i);
                const boldMatch = /\bBold=["'](?:true|1)["']/i.test(node) ||
                    (weightMatch && (/bold|black|heavy|semi[- ]?bold|demi[- ]?bold/i.test(weightMatch[1]) ||
                        parseInt(weightMatch[1], 10) >= 600));
                const italicMatch = node.match(/\bItalic="(true|1)"/i);

                if (idMatch && nameMatch) {
                    fontMap[idMatch[1]] = {
                        name: nameMatch[1],
                        bold: !!boldMatch,
                        italic: !!italicMatch
                    };
                }
            }
        }

        return fontMap;
    }

    // Extract Media Resources (Images) mapping
    extractMediaMap() {
        const mediaMap = {};
        const resFiles = [
            path.join(this.tempDir, 'Doc_0', 'PublicRes.xml'),
            path.join(this.tempDir, 'Doc_0', 'DocumentRes.xml')
        ];

        for (const resPath of resFiles) {
            if (!fs.existsSync(resPath)) continue;

            const xmlContent = fs.readFileSync(resPath, 'utf-8');
            const mediaNodes = xmlContent.match(/<(?:[a-zA-Z0-9]+:)?MultiMedia\b[\s\S]*?(?:\/>|><\/(?:[a-zA-Z0-9]+:)?MultiMedia>)/g) || [];

            for (const node of mediaNodes) {
                const idMatch = node.match(/\bID="([^"]+)"/) || node.match(/\bID='([^']+)'/);
                const locMatch = node.match(/\bBaseLoc="([^"]+)"/) || node.match(/\bBaseLoc='([^']+)'/) ||
                    node.match(/<(?:[a-zA-Z0-9]+:)?MediaFile>([^<]+)/);

                if (idMatch && locMatch) {
                    const relativeLoc = locMatch[1];
                    const fullPath = path.join(this.tempDir, 'Doc_0', relativeLoc);
                    if (fs.existsSync(fullPath)) {
                        mediaMap[idMatch[1]] = fullPath;
                    }
                }
            }
        }

        return mediaMap;
    }

    // Convert OFD color strings to Word Hex
    parseOfdColorToHex(colorStr) {
        if (!colorStr) return '000000';
        colorStr = colorStr.trim();

        if (colorStr.startsWith('#')) {
            return colorStr.replace('#', '').toUpperCase().padStart(6, '0').substring(0, 6);
        }

        const parts = colorStr.split(/[\s,]+/).map(Number).filter(n => !isNaN(n));
        if (parts.length >= 3) {
            let [r, g, b] = parts;
            if (r <= 1.0 && g <= 1.0 && b <= 1.0 && (r > 0 || g > 0 || b > 0)) {
                r = Math.round(r * 255);
                g = Math.round(g * 255);
                b = Math.round(b * 255);
            }
            const toHex = (n) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0').toUpperCase();
            return `${toHex(r)}${toHex(g)}${toHex(b)}`;
        }

        return '000000';
    }

    // OFD producers use a mixture of Boolean and CSS-like font-weight values.
    isBoldStyle(attributes) {
        const weightMatch = attributes.match(/\bWeight=["']([^"']+)["']/i);
        return /\bBold=["'](?:true|1)["']/i.test(attributes) ||
            !!(weightMatch && (/bold|black|heavy|semi[- ]?bold|demi[- ]?bold/i.test(weightMatch[1]) ||
                parseInt(weightMatch[1], 10) >= 600));
    }

    // Some OFD generators simulate bold by painting a filled glyph and then
    // painting its outline.  A stroke roughly 2% or more of the font size is
    // visibly heavier in Word, so map that representation to <w:b/>.
    isBoldStroke(attributes, fontSize) {
        const strokeMatch = attributes.match(/\bStroke=["'](?:true|1)["']/i);
        const widthMatch = attributes.match(/\bLineWidth=["']([^"']+)["']/i);
        const lineWidth = widthMatch ? parseFloat(widthMatch[1]) : 0;
        return !!strokeMatch && !isNaN(lineWidth) && lineWidth >= Math.max(0.1, fontSize * 0.02);
    }

    // Extract Text, Images, and Paths for a single page XML file
    extractPageElements(pageXmlPath, fontMap, mediaMap) {
        const elements = [];
        this.imagesToEmbed = this.imagesToEmbed || [];

        if (!fs.existsSync(pageXmlPath)) return elements;

        const xmlContent = fs.readFileSync(pageXmlPath, 'utf-8');

        // 1. Parse Text Objects
        const textObjectRegex = /<(?:[a-zA-Z0-9]+:)?TextObject\b([\s\S]*?)>([\s\S]*?)<\/(?:[a-zA-Z0-9]+:)?TextObject>/g;
        let match;
        while ((match = textObjectRegex.exec(xmlContent)) !== null) {
            const attributes = match[1];
            const body = match[2];

            const boundaryMatch = attributes.match(/Boundary="([^"]+)"/) || attributes.match(/Boundary='([^']+)'/);
            const fontMatch = attributes.match(/Font="([^"]+)"/) || attributes.match(/Font='([^']+)'/);
            const sizeMatch = attributes.match(/Size="([^"]+)"/) || attributes.match(/Size='([^']+)'/);
            const textCodes = [];
            const textCodeRegex = /<(?:[a-zA-Z0-9]+:)?TextCode\b([^>]*)>([\s\S]*?)<\/(?:[a-zA-Z0-9]+:)?TextCode>/g;
            let textCodeMatch;
            while ((textCodeMatch = textCodeRegex.exec(body)) !== null) {
                textCodes.push({ attributes: textCodeMatch[1], text: textCodeMatch[2] });
            }

            if (!boundaryMatch || !sizeMatch || textCodes.length === 0) continue;

            const [x, y, w, h] = boundaryMatch[1].split(' ').map(Number);
            const fontId = fontMatch ? fontMatch[1] : '';
            const sizeStr = sizeMatch[1];
            const text = textCodes.map(code => code.text).join('').trim();
            const fontSizePt = Math.round(parseFloat(sizeStr) * 2.83465);

            let colorStr = '';
            const attrColor = attributes.match(/FillColor="([^"]+)"/) || attributes.match(/FillColor='([^']+)'/);
            if (attrColor) {
                colorStr = attrColor[1];
            } else {
                const childColor = body.match(/<(?:[a-zA-Z0-9]+:)?FillColor\b[^>]*?\bValue="([^"]+)"/) ||
                    body.match(/<(?:[a-zA-Z0-9]+:)?FillColor\b[^>]*?\bValue='([^']+)'/);
                if (childColor) colorStr = childColor[1];
            }

            const fontObj = fontMap[fontId] || { name: 'SimSun', bold: false, italic: false };
            const isObjBold = this.isBoldStyle(attributes) || this.isBoldStroke(attributes, parseFloat(sizeStr));
            const isObjItalic = /\bItalic=["'](?:true|1)["']/i.test(attributes);
            const isUnderline = /\bUnderline=["'](?:true|1)["']/i.test(attributes);
            const isStrike = /\b(?:Strikeout|Strikethrough)=["'](?:true|1)["']/i.test(attributes);

            const alignMatch = attributes.match(/\bAlign="(left|center|right|justified)"/i);
            let align = alignMatch ? alignMatch[1].toLowerCase() : 'left';
            if (align === 'justified') align = 'both';

            const baseStyle = {
                fontName: fontObj.name,
                colorHex: this.parseOfdColorToHex(colorStr),
                fontSizeHalfPt: fontSizePt * 2,
                bold: isObjBold || fontObj.bold,
                italic: isObjItalic || fontObj.italic,
                underline: isUnderline,
                strike: isStrike
            };
            // TextCode normally only carries glyph positions, but some OFD
            // generators add style attributes there.  Keep each code as a Word
            // run so mixed bold/plain content survives either representation.
            const runs = textCodes.map(code => {
                const codeFontMatch = code.attributes.match(/\bFont=["']([^"']+)["']/i);
                const codeSizeMatch = code.attributes.match(/\bSize=["']([^"']+)["']/i);
                const codeColorMatch = code.attributes.match(/\bFillColor=["']([^"']+)["']/i);
                const codeFont = codeFontMatch ? (fontMap[codeFontMatch[1]] || fontObj) : fontObj;
                return {
                    text: code.text,
                    fontName: codeFont.name || baseStyle.fontName,
                    colorHex: codeColorMatch ? this.parseOfdColorToHex(codeColorMatch[1]) : baseStyle.colorHex,
                    fontSizeHalfPt: codeSizeMatch ? Math.round(parseFloat(codeSizeMatch[1]) * 2.83465) * 2 : baseStyle.fontSizeHalfPt,
                    bold: this.isBoldStyle(code.attributes) || this.isBoldStroke(code.attributes,
                        codeSizeMatch ? parseFloat(codeSizeMatch[1]) : parseFloat(sizeStr)) || baseStyle.bold || !!codeFont.bold,
                    italic: /\bItalic=["'](?:true|1)["']/i.test(code.attributes) || baseStyle.italic || !!codeFont.italic,
                    underline: /\bUnderline=["'](?:true|1)["']/i.test(code.attributes) || baseStyle.underline,
                    strike: /\b(?:Strikeout|Strikethrough)=["'](?:true|1)["']/i.test(code.attributes) || baseStyle.strike
                };
            }).filter(run => run.text.length > 0);

            elements.push({
                type: 'text',
                x, y, w, h,
                text,
                fontName: baseStyle.fontName,
                colorHex: baseStyle.colorHex,
                fontSizeHalfPt: baseStyle.fontSizeHalfPt,
                bold: baseStyle.bold,
                italic: baseStyle.italic,
                underline: baseStyle.underline,
                strike: baseStyle.strike,
                runs,
                align
            });
        }

        // 2. Parse Image Objects (Uses global this.imgIdx)
        const imageObjectRegex = /<(?:[a-zA-Z0-9]+:)?ImageObject\b([\s\S]*?)(?:\/>|><\/(?:[a-zA-Z0-9]+:)?ImageObject>)/g;
        while ((match = imageObjectRegex.exec(xmlContent)) !== null) {
            const attributes = match[1];
            const boundaryMatch = attributes.match(/Boundary="([^"]+)"/) || attributes.match(/Boundary='([^']+)'/);
            const resMatch = attributes.match(/ResourceID="([^"]+)"/) || attributes.match(/ResourceID='([^']+)'/);

            if (!boundaryMatch || !resMatch) continue;

            const [x, y, w, h] = boundaryMatch[1].split(' ').map(Number);
            const resourceId = resMatch[1];
            const imagePath = mediaMap[resourceId];

            if (imagePath && fs.existsSync(imagePath)) {
                const ext = path.extname(imagePath).toLowerCase() || '.png';
                const targetFileName = `image_${this.imgIdx}${ext}`;

                const cxEmu = Math.round(w * 36000);
                const cyEmu = Math.round(h * 36000);

                const imgData = {
                    type: 'image',
                    x, y, w, h,
                    sourcePath: imagePath,
                    targetFileName,
                    relId: `rIdImg${this.imgIdx}`,
                    cxEmu,
                    cyEmu,
                    imgIndex: this.imgIdx
                };

                elements.push(imgData);
                this.imagesToEmbed.push(imgData);
                this.imgIdx++;
            }
        }

        // 3. Parse Path Objects (Table Borders)
        // PathObject bodies commonly contain AbbreviatedData followed by
        // whitespace before the closing tag.  Capture attributes separately;
        // the former expression only matched a closing tag immediately after
        // '>' and therefore skipped normal table rules entirely.
        const pathObjectRegex = /<(?:[a-zA-Z0-9]+:)?PathObject\b([^>]*)(?:\/>|>[\s\S]*?<\/(?:[a-zA-Z0-9]+:)?PathObject>)/g;
        while ((match = pathObjectRegex.exec(xmlContent)) !== null) {
            const attributes = match[1];
            const boundaryMatch = attributes.match(/Boundary="([^"]+)"/) || attributes.match(/Boundary='([^']+)'/);

            if (!boundaryMatch) continue;

            const [x, y, w, h] = boundaryMatch[1].split(' ').map(Number);
            const isExplicitlyUnstroked = /\bStroke=["'](?:false|0)["']/i.test(attributes);
            const isFilled = /\bFill=["'](?:true|1)["']/i.test(attributes);

            // These are filled background rectangles used by some OFD writers
            // behind every visual text line.  They are neither table borders
            // nor paragraph separators; retaining them split a Chinese
            // paragraph into one Word paragraph per source line.
            if (isExplicitlyUnstroked && isFilled) continue;

            const strokeColorMatch = attributes.match(/StrokeColor="([^"]+)"/) || attributes.match(/StrokeColor='([^']+)'/);
            const strokeColor = strokeColorMatch ? this.parseOfdColorToHex(strokeColorMatch[1]) : 'CCCCCC';

            elements.push({
                type: 'path_border',
                x, y, w, h,
                // Table rules are normally long, thin PathObjects.  Retain
                // their orientation so they can later be reconstructed as a
                // native Word table rather than independent paragraphs.
                orientation: w >= h ? 'horizontal' : 'vertical',
                strokeColor
            });
        }

        return elements;
    }

    buildTableXml(table) {
        const escapeXml = (value) => String(value)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
        const renderRuns = (item) => (item.runs && item.runs.length ? item.runs : [item]).map(run => `
            <w:r><w:rPr>
                <w:rFonts w:ascii="${escapeXml(run.fontName)}" w:eastAsia="${escapeXml(run.fontName)}" w:hAnsi="${escapeXml(run.fontName)}" w:cs="${escapeXml(run.fontName)}"/>
                <w:color w:val="${run.colorHex}"/><w:sz w:val="${run.fontSizeHalfPt}"/>
                ${run.bold ? '<w:b/>' : ''}${run.italic ? '<w:i/>' : ''}
                ${run.underline ? '<w:u w:val="single"/>' : ''}${run.strike ? '<w:strike/>' : ''}
            </w:rPr><w:t xml:space="preserve">${escapeXml(run.text)}</w:t></w:r>`).join('');
        const borderColor = table.strokeColor || '000000';
        const borders = `<w:tblBorders><w:top w:val="single" w:sz="4" w:color="${borderColor}"/><w:left w:val="single" w:sz="4" w:color="${borderColor}"/><w:bottom w:val="single" w:sz="4" w:color="${borderColor}"/><w:right w:val="single" w:sz="4" w:color="${borderColor}"/><w:insideH w:val="single" w:sz="4" w:color="${borderColor}"/><w:insideV w:val="single" w:sz="4" w:color="${borderColor}"/></w:tblBorders>`;
        const grid = table.xs.slice(0, -1).map((x, i) => `<w:gridCol w:w="${Math.round((table.xs[i + 1] - x) * 56.7)}"/>`).join('');
        const rows = table.ys.slice(0, -1).map((y, row) => {
            const cells = table.xs.slice(0, -1).map((x, col) => {
                const right = table.xs[col + 1];
                const bottom = table.ys[row + 1];
                const cellItems = table.textItems.filter(item => {
                    const cx = item.x + item.w / 2;
                    const cy = item.y + item.h / 2;
                    return cx >= x && cx < right && cy >= y && cy < bottom;
                }).sort((a, b) => Math.abs(a.y - b.y) < 1.5 ? a.x - b.x : a.y - b.y);
                const textXml = cellItems.length ? cellItems.map((item, index) => {
                    const previous = index > 0 ? cellItems[index - 1] : null;
                    // A cell may contain multiple OFD visual lines.  Keep
                    // those as Word line breaks, while adjacent fragments on
                    // the same OFD line remain one continuous run.
                    const lineBreak = previous && Math.abs(item.y - previous.y) > 1.5
                        ? '<w:r><w:br/></w:r>' : '';
                    return lineBreak + renderRuns(item);
                }).join('') : '';
                return `<w:tc><w:tcPr><w:tcW w:w="${Math.round((right - x) * 56.7)}" w:type="dxa"/></w:tcPr><w:p>${textXml}</w:p></w:tc>`;
            }).join('');
            return `<w:tr>${cells}</w:tr>`;
        }).join('');
        return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${borders}</w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`;
    }

    reconstructTables(elements) {
        const rules = elements.filter(item => item.type === 'path_border');
        const horizontals = rules.filter(item => item.orientation === 'horizontal' && item.w > 5);
        const verticals = rules.filter(item => item.orientation === 'vertical' && item.h > 5);
        if (horizontals.length < 2 || verticals.length < 2) return elements;
        const unique = values => values.sort((a, b) => a - b).filter((value, i, all) => i === 0 || Math.abs(value - all[i - 1]) > 0.5);
        // Find a set of vertical rules sharing the same long y-span.  This is
        // the reliable table envelope; using every horizontal rule on a page
        // accidentally absorbed ordinary text above and below the table.
        const anchor = verticals.slice().sort((a, b) => b.h - a.h)[0];
        const overlap = (a, b) => Math.max(0, Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y));
        const gridVerticals = verticals.filter(item =>
            overlap(item, anchor) >= Math.min(item.h, anchor.h) * 0.8);
        if (gridVerticals.length < 2) return elements;

        const xs = unique(gridVerticals.map(item => item.x + item.w / 2));
        if (xs.length < 2) return elements;
        const left = xs[0], right = xs[xs.length - 1];
        const top = Math.min.apply(null, gridVerticals.map(item => item.y));
        const bottom = Math.max.apply(null, gridVerticals.map(item => item.y + item.h));
        const gridHorizontals = horizontals.filter(item => {
            const lineY = item.y + item.h / 2;
            return lineY >= top - 1 && lineY <= bottom + 1 &&
                item.x <= right - 1 && item.x + item.w >= left + 1;
        });
        const ys = unique(gridHorizontals.map(item => item.y + item.h / 2).concat([top, bottom]));
        if (gridHorizontals.length < 1 || ys.length < 2) return elements;
        const textItems = elements.filter(item => item.type === 'text' && item.x + item.w / 2 >= left && item.x + item.w / 2 < right && item.y + item.h / 2 >= top && item.y + item.h / 2 < bottom);
        const table = { xs, ys, textItems, strokeColor: gridHorizontals[0].strokeColor };
        const excluded = new Set(gridHorizontals.concat(gridVerticals, textItems));
        return elements.filter(item => !excluded.has(item)).concat({
            type: 'table', x: left, y: top, w: right - left, h: bottom - top,
            tableXml: this.buildTableXml(table)
        });
    }

    // Build Paragraph XML from layout elements
    buildParagraphs(elements) {
        elements = this.reconstructTables(elements);
        elements.sort((a, b) => (Math.abs(a.y - b.y) < 1.5 ? a.x - b.x : a.y - b.y));

        const lines = [];
        let currentLine = [];

        for (const el of elements) {
            if (currentLine.length === 0) {
                currentLine.push(el);
            } else {
                const prev = currentLine[currentLine.length - 1];
                if (el.type !== 'table' && prev.type !== 'table' && Math.abs(el.y - prev.y) < 2.0) {
                    currentLine.push(el);
                } else {
                    lines.push(currentLine);
                    currentLine = [el];
                }
            }
        }
        if (currentLine.length > 0) lines.push(currentLine);

        // OFD stores positioned text, not semantic paragraphs.  In particular,
        // Chinese documents commonly use the *same* line spacing inside and
        // between paragraphs.  Do not compare a gap with character height: the
        // OFD Boundary height is often not the actual line height.  Instead,
        // learn the normal line advance from this page and only split on a gap
        // that is clearly larger than the surrounding text.
        const paragraphs = [];
        let currentParagraph = [];

        const getTextMetrics = (line) => {
            const items = line.filter(item => item.type === 'text');
            if (items.length === 0) return null;
            return {
                top: Math.min.apply(null, items.map(item => item.y)),
                bottom: Math.max.apply(null, items.map(item => item.y + item.h)),
                height: Math.max.apply(null, items.map(item => item.h)),
                left: Math.min.apply(null, items.map(item => item.x))
            };
        };

        const textLines = lines.map(getTextMetrics).filter(Boolean);
        const lineAdvances = [];
        for (let i = 1; i < textLines.length; i++) {
            const advance = textLines[i].top - textLines[i - 1].top;
            if (advance > 0) lineAdvances.push(advance);
        }
        lineAdvances.sort((a, b) => a - b);
        const normalLineAdvance = lineAdvances.length > 0
            ? lineAdvances[Math.floor(lineAdvances.length / 2)]
            : 0;
        const pageTextLefts = textLines.map(item => item.left).sort((a, b) => a - b);
        // The median is resilient to page numbers and occasional indented
        // lines, while providing the body margin needed for a one-line heading.
        const pageBodyLeft = pageTextLefts.length > 0
            ? pageTextLefts[Math.floor(pageTextLefts.length / 2)]
            : 0;

        for (const line of lines) {
            const metrics = getTextMetrics(line);
            const previousLine = currentParagraph[currentParagraph.length - 1];
            const previousMetrics = previousLine && getTextMetrics(previousLine);

            if (!metrics || !previousMetrics) {
                if (currentParagraph.length > 0) paragraphs.push(currentParagraph);
                currentParagraph = [line];
                continue;
            }

            // Chinese first-line indentation is normally two characters.  It
            // is useful evidence of a paragraph boundary even when there is no
            // extra vertical space.  Require more than one character to avoid
            // treating small coordinate rounding differences as indentation.
            // Compare consecutive lines, not the page-wide left edge.  A page
            // number, heading, or table can have a much smaller x-coordinate
            // than body text; using that global value made every normal body
            // line appear to be a new paragraph.
            const firstLineIndent = metrics.left - previousMetrics.left;
            const indentThreshold = Math.max(3, metrics.height * 1.2);
            const hasFirstLineIndent = firstLineIndent >= indentThreshold;
            const lineAdvance = metrics.top - previousMetrics.top;
            const hasLargeVerticalGap = normalLineAdvance > 0 &&
                lineAdvance > normalLineAdvance * 1.45;
            const previousTextItems = previousLine.filter(item => item.type === 'text');
            const currentTextItems = line.filter(item => item.type === 'text');
            // A short, bold-only line is a section heading.  This separates
            // "（二）职责分工" from the indented body paragraph without
            // splitting a line that merely begins with bold inline text.
            const hasHeadingBreak = previousTextItems.length > 0 &&
                previousTextItems.every(item => item.bold) &&
                previousTextItems.reduce((length, item) => length + item.text.length, 0) <= 30 &&
                currentTextItems.some(item => !item.bold);

            if (hasFirstLineIndent || hasLargeVerticalGap || hasHeadingBreak) {
                paragraphs.push(currentParagraph);
                currentParagraph = [line];
            } else {
                currentParagraph.push(line);
            }
        }
        if (currentParagraph.length > 0) paragraphs.push(currentParagraph);

        const escapeXml = (value) => String(value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');

        return paragraphs.map(paragraph => {
            const textItems = [];
            const imageItems = [];
            const pathItems = [];
            const tableItems = [];
            paragraph.forEach((line, lineIndex) => {
                line.forEach(item => {
                    if (item.type === 'text') textItems.push({ item, lineIndex });
                    else if (item.type === 'image') imageItems.push(item);
                    else if (item.type === 'path_border') pathItems.push(item);
                    else if (item.type === 'table') tableItems.push(item);
                });
            });

            let paragraphContent = '';

            if (tableItems.length > 0) return tableItems.map(item => item.tableXml).join('');

            if (textItems.length > 0) {
                const firstItem = textItems[0].item;
                const lineAlign = firstItem.align || 'left';
                // Derive indentation from this paragraph's own lines.  Do not
                // use the page-wide leftmost object, which may be a page number
                // or a table cell unrelated to this paragraph.
                const paragraphBodyLeft = Math.min.apply(null, textItems.map(entry => entry.item.x));
                const paragraphIndent = Math.max(0, firstItem.x - Math.min(paragraphBodyLeft, pageBodyLeft));
                const firstLineIndentXml = paragraphIndent >= Math.max(3, firstItem.h * 1.2)
                    ? `<w:ind w:firstLine="${Math.round(paragraphIndent * 56.7)}"/>`
                    : '';
                const jcXml = lineAlign !== 'left' || firstLineIndentXml
                    ? `<w:pPr>${firstLineIndentXml}${lineAlign !== 'left' ? `<w:jc w:val="${lineAlign}"/>` : ''}</w:pPr>`
                    : '';

                const runsXml = textItems.map((entry, index) => {
                    const item = entry.item;
                    const previousItem = index > 0 ? textItems[index - 1].item : null;
                    // OFD splits a paragraph into physical lines.  Preserve an
                    // explicit space when present; otherwise insert one only
                    // between two ASCII words.  Chinese text needs no space.
                    const needsSpace = previousItem && entry.lineIndex !== textItems[index - 1].lineIndex &&
                        /[A-Za-z0-9]$/.test(previousItem.text) && /^[A-Za-z0-9]/.test(item.text);
                    const itemRuns = item.runs && item.runs.length > 0 ? item.runs : [item];
                    const itemRunsXml = itemRuns.map(run => `
                    <w:r>
                        <w:rPr>
                            <w:rFonts w:ascii="${escapeXml(run.fontName)}" w:eastAsia="${escapeXml(run.fontName)}" w:hAnsi="${escapeXml(run.fontName)}" w:cs="${escapeXml(run.fontName)}" />
                            <w:color w:val="${run.colorHex}" />
                            <w:sz w:val="${run.fontSizeHalfPt}" />
                            ${run.bold ? '<w:b/>' : ''}
                            ${run.italic ? '<w:i/>' : ''}
                            ${run.underline ? '<w:u w:val="single"/>' : ''}
                            ${run.strike ? '<w:strike/>' : ''}
                        </w:rPr>
                        <w:t xml:space="preserve">${escapeXml(run.text)}</w:t>
                    </w:r>`).join('');
                    return `${needsSpace ? '<w:r><w:t xml:space="preserve"> </w:t></w:r>' : ''}${itemRunsXml}`;
                }).join('');

                paragraphContent += `${jcXml}${runsXml}`;
            }

            if (imageItems.length > 0) {
                const imgXml = imageItems.map(img => `
                    <w:r>
                        <w:drawing>
                            <wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
                                <wp:extent cx="${img.cxEmu}" cy="${img.cyEmu}"/>
                                <wp:docPr id="${img.imgIndex}" name="Picture ${img.imgIndex}"/>
                                <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                            <pic:nvPicPr>
                                                <pic:cNvPr id="${img.imgIndex}" name="Picture ${img.imgIndex}"/>
                                                <pic:cNvPicPr/>
                                            </pic:nvPicPr>
                                            <pic:blipFill>
                                                <a:blip r:embed="${img.relId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                                                <a:stretch><a:fillRect/></a:stretch>
                                            </pic:blipFill>
                                            <pic:spPr>
                                                <a:xfrm><a:off x="0" y="0"/><a:ext cx="${img.cxEmu}" cy="${img.cyEmu}"/></a:xfrm>
                                                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                                            </pic:spPr>
                                        </pic:pic>
                                    </a:graphicData>
                                </a:graphic>
                            </wp:inline>
                        </w:drawing>
                    </w:r>`).join('');

                paragraphContent += imgXml;
            }

            // A PathObject is a drawing primitive, not a Word paragraph.
            // When it does not belong to a reconstructed table, omit it rather
            // than generating a stray paragraph bottom border between text.
            if (pathItems.length > 0 && textItems.length === 0 && imageItems.length === 0) return '';

            return `<w:p>${paragraphContent}</w:p>`;
        }).join('');
    }

    // Pack generated DOCX structure
    generateDocxPackage(paragraphsXml) {
        const wordSubDir = path.join(this.docxDir, 'word');
        const mediaSubDir = path.join(wordSubDir, 'media');
        const relsSubDir = path.join(this.docxDir, '_rels');
        const wordRelsSubDir = path.join(wordSubDir, '_rels');

        fs.mkdirSync(wordSubDir, { recursive: true });
        fs.mkdirSync(relsSubDir, { recursive: true });
        fs.mkdirSync(wordRelsSubDir, { recursive: true });

        const images = this.imagesToEmbed || [];

        if (images.length > 0) {
            fs.mkdirSync(mediaSubDir, { recursive: true });
            for (const img of images) {
                const targetPath = path.join(mediaSubDir, img.targetFileName);
                fs.copyFileSync(img.sourcePath, targetPath);
            }
        }

        // 1. [Content_Types].xml
        fs.writeFileSync(path.join(this.docxDir, '[Content_Types].xml'),
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Default Extension="png" ContentType="image/png"/>
            <Default Extension="jpg" ContentType="image/jpeg"/>
            <Default Extension="jpeg" ContentType="image/jpeg"/>
            <Default Extension="bmp" ContentType="image/bmp"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>`);

        // 2. _rels/.rels
        fs.writeFileSync(path.join(relsSubDir, '.rels'),
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>`);

        // 3. word/_rels/document.xml.rels
        const imageRelsXml = images.map(img => `
            <Relationship Id="${img.relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${img.targetFileName}"/>
        `).join('');

        fs.writeFileSync(path.join(wordRelsSubDir, 'document.xml.rels'),
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            ${imageRelsXml}
        </Relationships>`);

        // 4. word/document.xml
        fs.writeFileSync(path.join(wordSubDir, 'document.xml'),
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                ${paragraphsXml}
            </w:body>
        </w:document>`);

        // Compress directory to .docx
        if (process.platform === 'win32') {
            const tempOutZip = path.join(process.cwd(), `temp_out_${Date.now()}.zip`);
            const emptyZipHeader = Buffer.from([80, 75, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
            fs.writeFileSync(tempOutZip, emptyZipHeader);

            const winZip = path.resolve(tempOutZip).replace(/\//g, '\\');
            const winSrc = path.resolve(this.docxDir).replace(/\//g, '\\');

            try {
                const psScript = `$s = New-Object -ComObject Shell.Application; $z = $s.NameSpace('${winZip}'); $d = $s.NameSpace('${winSrc}'); if(-not $z -or -not $d){ exit 1 }; $items = $z.Items(); $z.CopyHere($items); $start = Get-Date; while(($z.Items().Count -lt $items.Count) -and ((Get-Date) - $start).TotalSeconds -lt 10){ Start-Sleep -m 100 }`;
                execSync(`powershell -NoProfile -NonInteractive -Command "${psScript}"`);

                if (fs.existsSync(this.outputDocxPath)) {
                    fs.unlinkSync(this.outputDocxPath);
                }
                fs.renameSync(tempOutZip, this.outputDocxPath);
            } catch (e) {
                if (fs.existsSync(tempOutZip)) fs.unlinkSync(tempOutZip);
                throw e;
            }
        } else {
            execSync(`cd "${this.docxDir}" && zip -q -r "${this.outputDocxPath}" ./*`);
        }
    }

    removeDirSync(dirPath) {
        if (fs.existsSync(dirPath)) {
            fs.rmdirSync(dirPath, { recursive: true });
        }
    }

    cleanup() {
        this.removeDirSync(this.tempDir);
        this.removeDirSync(this.docxDir);
    }

    // Primary Conversion Pipeline
    convert() {
        try {
            this.imagesToEmbed = [];
            this.imgIdx = 1;

            this.unzipOFD();

            const fontMap = this.extractFontMap();
            const mediaMap = this.extractMediaMap();
            const pagePaths = this.getPagePaths();

            if (pagePaths.length === 0) {
                console.warn('No OFD pages found to convert.');
                return;
            }

            let fullDocumentXml = '';

            for (let i = 0; i < pagePaths.length; i++) {
                const pagePath = pagePaths[i];
                const elements = this.extractPageElements(pagePath, fontMap, mediaMap);
                const pageParagraphsXml = this.buildParagraphs(elements);

                // Add OpenXML hard page break between pages
                if (i > 0) {
                    fullDocumentXml += `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
                }

                fullDocumentXml += pageParagraphsXml;
            }

            this.generateDocxPackage(fullDocumentXml);
            console.log(`Successfully converted ${pagePaths.length} page(s) to: ${this.outputDocxPath}`);
        } catch (err) {
            console.error('Error converting multi-page file:', err);
        } finally {
            this.cleanup();
        }
    }
}

// Execution
const converter = new OFD2WordConverter('./sample.ofd', './output.docx');
converter.convert();
