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
                const boldMatch = node.match(/\bBold="(true|1)"/i) || node.match(/\bWeight="(bold|700)"/i);
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
            const textCodeRegex = /<(?:[a-zA-Z0-9]+:)?TextCode[^>]*>([\s\S]*?)<\/(?:[a-zA-Z0-9]+:)?TextCode>/g;
            let textCodeMatch;
            while ((textCodeMatch = textCodeRegex.exec(body)) !== null) {
                textCodes.push(textCodeMatch[1]);
            }

            if (!boundaryMatch || !sizeMatch || textCodes.length === 0) continue;

            const [x, y, w, h] = boundaryMatch[1].split(' ').map(Number);
            const fontId = fontMatch ? fontMatch[1] : '';
            const sizeStr = sizeMatch[1];
            const text = textCodes.join('').trim();
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
            const isObjBold = /\bBold="(true|1)"/i.test(attributes) || /\bWeight="(bold|700)"/i.test(attributes);
            const isObjItalic = /\bItalic="(true|1)"/i.test(attributes);
            const isUnderline = /\bUnderline="(true|1)"/i.test(attributes);
            const isStrike = /\b(Strikeout|Strikethrough)="(true|1)"/i.test(attributes);

            const alignMatch = attributes.match(/\bAlign="(left|center|right|justified)"/i);
            let align = alignMatch ? alignMatch[1].toLowerCase() : 'left';
            if (align === 'justified') align = 'both';

            elements.push({
                type: 'text',
                x, y, w, h,
                text,
                fontName: fontObj.name,
                colorHex: this.parseOfdColorToHex(colorStr),
                fontSizeHalfPt: fontSizePt * 2,
                bold: isObjBold || fontObj.bold,
                italic: isObjItalic || fontObj.italic,
                underline: isUnderline,
                strike: isStrike,
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
        const pathObjectRegex = /<(?:[a-zA-Z0-9]+:)?PathObject\b([\s\S]*?)(?:\/>|><\/(?:[a-zA-Z0-9]+:)?PathObject>)/g;
        while ((match = pathObjectRegex.exec(xmlContent)) !== null) {
            const attributes = match[1];
            const boundaryMatch = attributes.match(/Boundary="([^"]+)"/) || attributes.match(/Boundary='([^']+)'/);

            if (!boundaryMatch) continue;

            const [x, y, w, h] = boundaryMatch[1].split(' ').map(Number);
            const strokeColorMatch = attributes.match(/StrokeColor="([^"]+)"/) || attributes.match(/StrokeColor='([^']+)'/);
            const strokeColor = strokeColorMatch ? this.parseOfdColorToHex(strokeColorMatch[1]) : 'CCCCCC';

            elements.push({
                type: 'path_border',
                x, y, w, h,
                strokeColor
            });
        }

        return elements;
    }

    // Build Paragraph XML from layout elements
    buildParagraphs(elements) {
        elements.sort((a, b) => (Math.abs(a.y - b.y) < 1.5 ? a.x - b.x : a.y - b.y));

        const lines = [];
        let currentLine = [];

        for (const el of elements) {
            if (currentLine.length === 0) {
                currentLine.push(el);
            } else {
                const prev = currentLine[currentLine.length - 1];
                if (Math.abs(el.y - prev.y) < 2.0) {
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
        const bodyLefts = textLines.map(item => item.left).sort((a, b) => a - b);
        const normalLeft = bodyLefts.length > 0
            ? bodyLefts[Math.floor(bodyLefts.length / 2)]
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

            const lineAdvance = metrics.top - previousMetrics.top;
            // Chinese first-line indentation is normally two characters.  It
            // is useful evidence of a paragraph boundary even when there is no
            // extra vertical space.  Require more than one character to avoid
            // treating small coordinate rounding differences as indentation.
            const firstLineIndent = metrics.left - normalLeft;
            const indentThreshold = Math.max(3, metrics.height * 1.2);
            const hasFirstLineIndent = firstLineIndent >= indentThreshold;
            const hasLargeVerticalGap = normalLineAdvance > 0 &&
                lineAdvance > normalLineAdvance * 1.45;

            if (hasFirstLineIndent || hasLargeVerticalGap) {
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
            paragraph.forEach((line, lineIndex) => {
                line.forEach(item => {
                    if (item.type === 'text') textItems.push({ item, lineIndex });
                    else if (item.type === 'image') imageItems.push(item);
                    else if (item.type === 'path_border') pathItems.push(item);
                });
            });

            let paragraphContent = '';

            if (textItems.length > 0) {
                const firstItem = textItems[0].item;
                const lineAlign = firstItem.align || 'left';
                const paragraphIndent = Math.max(0, firstItem.x - normalLeft);
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
                    return `${needsSpace ? '<w:r><w:t xml:space="preserve"> </w:t></w:r>' : ''}
                    <w:r>
                        <w:rPr>
                            <w:rFonts w:ascii="${escapeXml(item.fontName)}" w:eastAsia="${escapeXml(item.fontName)}" w:hAnsi="${escapeXml(item.fontName)}" w:cs="${escapeXml(item.fontName)}" />
                            <w:color w:val="${item.colorHex}" />
                            <w:sz w:val="${item.fontSizeHalfPt}" />
                            ${item.bold ? '<w:b/>' : ''}
                            ${item.italic ? '<w:i/>' : ''}
                            ${item.underline ? '<w:u w:val="single"/>' : ''}
                            ${item.strike ? '<w:strike/>' : ''}
                        </w:rPr>
                        <w:t xml:space="preserve">${escapeXml(item.text)}</w:t>
                    </w:r>`;
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

            if (pathItems.length > 0 && textItems.length === 0 && imageItems.length === 0) {
                const borderItem = pathItems[0];
                return `<w:p>
                    <w:pPr>
                        <w:pBdr>
                            <w:bottom w:val="single" w:sz="6" w:space="1" w:color="${borderItem.strokeColor}"/>
                        </w:pBdr>
                    </w:pPr>
                </w:p>`;
            }

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
