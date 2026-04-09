const Docxtemplater = require('docxtemplater');
const PizZip = require('pizzip');

const zip = new PizZip();
// Create a fake docx structure for testing
zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Section 184(1) and Rule 9(2). Name: (name of the Company)</w:t></w:r></w:p></w:body></w:document>`);

zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`);

zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`);

try {
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        delimiters: { start: '(', end: ')' },
        nullGetter(part) {
            if (!part.module) {
                return "(" + part.value + ")";
            }
            if (part.module === "rawxml") {
                return "";
            }
            return "";
        }
    });

    doc.render({
        "name of the Company": "TestCompany"
    });

    const out = doc.getZip().file("word/document.xml").asText();
    console.log(out);
} catch (e) {
    console.error(e);
}
