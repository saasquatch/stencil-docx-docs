import * as fs from "fs";
import * as docx from "docx";
import {
  JsonDocs,
  JsonDocsComponent,
  JsonDocsProp,
  JsonDocsSlot,
} from "@stencil/core/internal";

export interface GenerateDocxOptions {
  /** The output directory for the docx */
  outDir?: string;
  /** The output file name (.docx) */
  outFile?: string;
  /** The name of the font used for the text in the document */
  textFont?: string;
  /** Doc tags on components or props that would cause them to be excluded from the output */
  excludeTags?: string[];
  /** Title text for the title page */
  title?: string;
  /** Author for the title page */
  author?: string;
}

const defaultOptions: GenerateDocxOptions = {
  outDir: "docs",
  outFile: "docs.docx",
  textFont: "Calibri",
  excludeTags: ["undocumented"],
  title: "Component Documentation",
  author: "SaaSquatch",
};

//const portraitPageWidth = 9638;
const landscapePageWidth = 13768;

const propTableColumnWidths = [
  landscapePageWidth * 0.25,
  landscapePageWidth * 0.15,
  landscapePageWidth * 0.6,
];

const propTableHeaderCells = () => [
  makeHeaderCell("Attribute Name"),
  makeHeaderCell("Type"),
  makeHeaderCell("Description"),
];

const propToCells = (prop: JsonDocsProp) => [
  makeCell(prop.attr || prop.name),
  makeCell(prop.type),
  makeCell(prop.docs),
];

const slotTableColumnWidths = [
  landscapePageWidth * 0.2,
  landscapePageWidth * 0.8,
];

const slotTableHeaderCells = () => [
  makeHeaderCell("Name"),
  makeHeaderCell("Description"),
];

const slotToCells = (slot: JsonDocsSlot) => [
  makeCell(slot.name),
  makeCell(slot.docs),
];

function makePageBreak() {
  return new docx.Paragraph({
    children: [new docx.PageBreak()],
  });
}

function makeHeaderCell(text: string) {
  return new docx.TableCell({
    margins: {
      top: 64,
      bottom: 64,
      left: 64,
      right: 64,
    },
    shading: {
      color: "auto",
      fill: "eeeeee",
      val: docx.ShadingType.CLEAR,
    },
    children: [
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text,
            bold: true,
          }),
        ],
      }),
    ],
  });
}

function makeCell(text: string) {
  return new docx.TableCell({
    margins: {
      top: 64,
      bottom: 64,
      left: 64,
      right: 64,
    },
    children: [
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text,
          }),
        ],
      }),
    ],
  });
}

function makeParagraph(
  text: string,
  paragraphOptions?: docx.IParagraphOptions
) {
  return new docx.Paragraph({
    ...paragraphOptions,
    children: [
      new docx.TextRun({
        text,
      }),
    ],
  });
}

function makeComponent(
  component: JsonDocsComponent,
  options: GenerateDocxOptions
) {
  const propTable = component.props.length
    ? makePropTable(component.props, options)
    : null;

  const slotTable = component.slots.length
    ? makeSlotTable(component.slots)
    : null;

  const children: docx.ISectionOptions["children"] = [
    makeParagraph(component.tag, { heading: docx.HeadingLevel.HEADING_1 }),
    makeParagraph(
      component.docs || "No top-level documentation for this component."
    ),
  ];

  if (propTable) {
    children.push(
      makeParagraph("Props", { heading: docx.HeadingLevel.HEADING_2 })
    );
    children.push(propTable);
  }

  if (slotTable) {
    children.push(
      makeParagraph("Slots", { heading: docx.HeadingLevel.HEADING_2 })
    );
    children.push(slotTable);
  }

  return children;
}

function makePropTable(props: JsonDocsProp[], options: GenerateDocxOptions) {
  const rows = [
    new docx.TableRow({
      tableHeader: true,
      children: propTableHeaderCells(),
    }),
    ...props
      .filter(
        (prop) =>
          !prop.docsTags.some((tag) => options.excludeTags.includes(tag.name))
      )
      .map(
        (prop) =>
          new docx.TableRow({
            cantSplit: false,
            children: propToCells(prop),
          })
      ),
  ];

  const table = new docx.Table({
    columnWidths: propTableColumnWidths,
    rows,
  });

  return table;
}

function makeSlotTable(slots: JsonDocsSlot[]) {
  const rows = [
    new docx.TableRow({
      tableHeader: true,
      children: slotTableHeaderCells(),
    }),
    ...slots.map(
      (slot) =>
        new docx.TableRow({
          cantSplit: false,
          children: slotToCells(slot),
        })
    ),
  ];

  const table = new docx.Table({
    columnWidths: slotTableColumnWidths,
    rows,
  });

  return table;
}

function generateDocx(options: GenerateDocxOptions, docs: JsonDocs) {
  const doc = new docx.Document({
    creator: options.author,
    title: options.title,
    styles: {
      paragraphStyles: [
        {
          id: "Normal",
          name: "Normal",
          quickFormat: true,
          run: {
            size: 24,
            font: {
              name: options.textFont,
            },
          },
        },
        {
          id: "Heading1",
          name: "Heading 1",
          quickFormat: true,
          paragraph: {
            outlineLevel: 1,
            spacing: {
              before: 240,
              after: 120,
            },
          },
          run: {
            size: 32,
            bold: true,
            font: {
              name: options.textFont,
            },
          },
        },
        {
          id: "Heading2",
          name: "Heading 2",
          quickFormat: true,
          paragraph: {
            outlineLevel: 2,
            spacing: {
              before: 240,
              after: 120,
            },
          },
          run: {
            size: 28,
            bold: true,
            font: {
              name: options.textFont,
            },
          },
        },
        {
          id: "Title",
          name: "Title",
          quickFormat: true,
          paragraph: {
            spacing: {
              before: 480,
              after: 240,
            },
          },
          run: {
            size: 64,
            bold: true,
            font: {
              name: options.textFont,
            },
          },
        },
        {
          id: "Subtitle",
          name: "Subtitle",
          quickFormat: true,
          paragraph: {
            spacing: {
              before: 360,
              after: 180,
            },
          },
          run: {
            size: 36,
            bold: true,
            font: {
              name: options.textFont,
            },
          },
        },
        {
          id: "Heading1NoOutline",
          name: "Heading 1 No Outline",
          quickFormat: false,
          paragraph: {
            spacing: {
              before: 240,
              after: 120,
            },
          },
          run: {
            size: 32,
            bold: true,
            font: {
              name: options.textFont,
            },
          },
        },
      ],
    },
  });

  doc.addSection({
    children: [
      makeParagraph(options.title, { style: "Title" }),
      makeParagraph(
        `Generated by ${options.author} on ${new Date().toLocaleDateString()}`,
        { style: "Subtitle" }
      ),
    ],
  });

  const components = docs.components
    .filter(
      (component) =>
        !component.docsTags.some((tag) =>
          options.excludeTags.includes(tag.name)
        )
    )
    .reduce(
      (children, component) =>
        children.concat(makeComponent(component, options)),
      [] as docx.ISectionOptions["children"]
    );

  doc.addSection({
    children: [
      makeParagraph("Table of Contents", {
        style: "Heading1NoOutline",
      }),
      new docx.TableOfContents("Table of Contents", {
        hyperlink: true,
        headingStyleRange: "1-2",
      }),
      makePageBreak(),
      ...components,
    ],

    properties: {
      orientation: docx.PageOrientation.LANDSCAPE,
    },

    footers: {
      default: new docx.Footer({
        children: [
          new docx.Table({
            width: {
              size: landscapePageWidth,
              type: docx.WidthType.DXA,
            },
            rows: [
              new docx.TableRow({
                children: [
                  new docx.TableCell({
                    borders: {
                      top: {
                        size: 0,
                        style: docx.BorderStyle.NONE,
                        color: "0",
                      },
                      bottom: {
                        size: 0,
                        style: docx.BorderStyle.NONE,
                        color: "0",
                      },
                      left: {
                        size: 0,
                        style: docx.BorderStyle.NONE,
                        color: "0",
                      },
                      right: {
                        size: 0,
                        style: docx.BorderStyle.NONE,
                        color: "0",
                      },
                    },
                    children: [makeParagraph(options.title)],
                  }),
                  new docx.TableCell({
                    borders: {
                      top: {
                        size: 0,
                        style: docx.BorderStyle.NONE,
                        color: "0",
                      },
                      bottom: {
                        size: 0,
                        style: docx.BorderStyle.NONE,
                        color: "0",
                      },
                      left: {
                        size: 0,
                        style: docx.BorderStyle.NONE,
                        color: "0",
                      },
                      right: {
                        size: 0,
                        style: docx.BorderStyle.NONE,
                        color: "0",
                      },
                    },
                    children: [
                      new docx.Paragraph({
                        alignment: docx.AlignmentType.RIGHT,
                        children: [
                          new docx.TextRun({
                            children: [docx.PageNumber.CURRENT],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    },
  });

  if (!fs.existsSync(options.outDir)) {
    fs.mkdirSync(options.outDir);
  }

  const outFile = `${options.outDir}/${options.outFile}`;
  console.log(`Writing .docx component documentation to ${outFile}...`);

  docx.Packer.toBuffer(doc).then((buffer) => fs.writeFileSync(outFile, buffer));
}

function createDocxGenerator(options?: GenerateDocxOptions) {
  return generateDocx.bind(this, { ...defaultOptions, ...options });
}

export default createDocxGenerator;
