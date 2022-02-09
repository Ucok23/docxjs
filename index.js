const docx = require("docx");
const express = require("express");
const path = require('path');
const app = express();

const port = 3000;
const { Document, 
        HeadingLevel, 
        Packer, 
        Paragraph, 
        Table, TableCell, TableRow,
        TextDirection,
        VerticalAlign,
        WidthType } = docx;

class DocumentCreator {
    create() {
        const doc = new Document({
            sections: [{
                children: [
                    this.createHeading1('This Is The Heading1'),
                    new Paragraph({}),
                    this.createHeading2('This Is The Heading2'),
                    new Paragraph({}),
                    this.createHeading3('This Is The Heading3'),
                    new Paragraph({}),
                    new Paragraph('Normal paragraph'),
                    new Paragraph({}),
                    this.createTable()
                ]
            }]
        });
        return doc;
    }

    createHeading1(text) {
        return new Paragraph({
            text: text,
            heading: HeadingLevel.HEADING_1,
            thematicBreak: true,
        });
    }

    createHeading2(text) {
        return new Paragraph({
            text: text,
            heading: HeadingLevel.HEADING_2,
            thematicBreak: true,
        });
    }

    createHeading3(text) {
        return new Paragraph({
            text: text,
            heading: HeadingLevel.HEADING_3,
            thematicBreak: true,
        });
    }

    createTable() {
        const table = new Table({
            rows: [
                new TableRow({
                    children: [
                        new TableCell({
                            children: [new Paragraph('First')],
                            verticalAlign: VerticalAlign.CENTER,
                        }),
                        new TableCell({
                            children: [new Paragraph('Second')],
                            verticalAlign: VerticalAlign.CENTER,
                        }),
                        new TableCell({
                            children: [new Paragraph({ text: "Third" })],
                            textDirection: TextDirection.CENTER,
                        }),
                        new TableCell({
                            children: [new Paragraph({ text: "Fourth" })],
                            textDirection: TextDirection.CENTER,
                        }),
                    ],
                }),
                new TableRow({
                    children: [
                        new TableCell({
                            children: [new Paragraph({text:'Fifth'})],
                            textDirection: TextDirection.CENTER,
                        }),
                        new TableCell({
                            children: [
                                new Paragraph({
                                    text: "Sixth",
                                }),
                            ],
                            textDirection: TextDirection.CENTER,
                        }),
                        new TableCell({
                            children: [
                                new Paragraph({
                                    text: "Seventh",
                                }),
                            ],
                            textDirection: TextDirection.CENTER,
                        }),
                        new TableCell({
                            children: [
                                new Paragraph({
                                    text: "Eighth",
                                }),
                            ],
                            textDirection: TextDirection.CENTER,
                        }),
                    ],
                }),
            ],
            width: {
                size: 4535,
                type: WidthType.DXA
            }
        })
        return table;
    }
}

app.get("/", async (req, res) => {
    res.sendFile(path.join(__dirname, './index.html'));
});

app.get("/download", async (req, res) => {
    const filename = 'Simple Doc';
    const documentCreator = new DocumentCreator();
    const doc = documentCreator.create();

    const b64string = await Packer.toBase64String(doc);
    
    res.setHeader('Content-Disposition', `attachment; filename=${filename}.docx`);
    res.send(Buffer.from(b64string, 'base64'));
});

app.listen(port);