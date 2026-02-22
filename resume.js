const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, LevelFormat,
  HeadingLevel, UnderlineType
} = require('docx');
const fs = require('fs');
const path = require('path');

const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

const sectionHeading = (text) => new Paragraph({
  children: [new TextRun({ text, bold: true, size: 26, color: "2E75B6" })],
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 1 } },
  spacing: { before: 240, after: 80 }
});

const jobTitle = (title, company) => new Paragraph({
  children: [
    new TextRun({ text: title, bold: true, size: 22 }),
    new TextRun({ text: " | ", bold: true, size: 22 }),
    new TextRun({ text: company, bold: true, size: 22 })
  ],
  spacing: { before: 120, after: 40 }
});

const jobMeta = (meta) => new Paragraph({
  children: [new TextRun({ text: meta, size: 18, color: "666666", italics: true })],
  spacing: { after: 60 }
});

const bulletParagraph = (text) => ({
  numbering: { reference: "bullets", level: 0 },
  children: [new TextRun({ text, size: 20 })],
  spacing: { after: 60 }
});

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0,
        format: LevelFormat.BULLET,
        text: "\u2022",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } }
      }]
    }]
  },
  styles: {
    default: {
      document: { run: { font: "Calibri", size: 20 } }
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    children: [
      // Name
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Alexey Folomeev", bold: true, size: 52, font: "Calibri" })],
        spacing: { after: 60 }
      }),
      // Title
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Fullstack Software Developer", size: 32, color: "444444" })],
        spacing: { after: 80 }
      }),
      // Contact
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "folomeev.aleksei97@gmail.com  |  +996-508-25-05-97  |  github.com/VisionFoland2025", size: 18, color: "555555" })],
        spacing: { after: 80 }
      }),
      // Summary
      new Paragraph({
        children: [new TextRun({
          text: "Enthusiastic Fullstack Engineer specializing in Java/Spring and modern JavaScript. I thrive on solving intricate problems and streamlining development processes. Committed to technical excellence and timely project delivery, I’m looking to leverage my coding skills to build robust software solutions in a fast-paced entry-level role.",
          size: 20
        })],
        spacing: { after: 200 }
      }),

      // Two-column layout: Skills | Experience
      new Table({
        width: { size: 10080, type: WidthType.DXA },
        columnWidths: [3000, 7080],
        rows: [
          new TableRow({
            children: [
              // LEFT: Skills + Education + Languages
              new TableCell({
                borders: { ...noBorders, right: { style: BorderStyle.SINGLE, size: 4, color: "DDDDDD" } },
                width: { size: 3000, type: WidthType.DXA },
                margins: { top: 0, bottom: 0, left: 0, right: 200 },
                children: [
                  // SKILLS
                  new Paragraph({
                    children: [new TextRun({ text: "Skills", bold: true, size: 26, color: "2E75B6" })],
                    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 1 } },
                    spacing: { after: 120 }
                  }),
                  ...[
                    { skill: "Javascript / React / RTK / NextJS / TypeScript", stars: 5 },
                    { skill: "HTML / CSS3 / SASS / TailwindCSS", stars: 5 },
                    { skill: "Java / Spring Boot", stars: 5 },
                    { skill: "Oracle SQL and PL/SQL / PostgreSQL / Clickhouse", stars: 5 },
                    { skill: "Git / Docker / Linux", stars: 5 }
                  ].flatMap(({ skill, stars }) => [
                    new Paragraph({
                      children: [new TextRun({ text: skill, bold: true, size: 20 })],
                      spacing: { before: 100, after: 20 }
                    }),
                    new Paragraph({
                      children: [new TextRun({
                        text: "★".repeat(stars) + "☆".repeat(5 - stars),
                        size: 22,
                        color: "F5A623"
                      })],
                      spacing: { after: 80 }
                    })
                  ]),

                  // EDUCATION
                  new Paragraph({
                    children: [new TextRun({ text: "Education", bold: true, size: 26, color: "2E75B6" })],
                    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 1 } },
                    spacing: { before: 240, after: 120 }
                  }),
                  new Paragraph({
                    children: [new TextRun({ text: "INRTU", bold: true, size: 18 })],
                    spacing: { after: 30 }
                  }),
                  new Paragraph({
                    children: [new TextRun({ text: "July 2014 – August 2020", size: 18, color: "666666" })],
                    spacing: { after: 120 }
                  }),

                  // COURSES
                  new Paragraph({
                    children: [new TextRun({ text: "Courses", bold: true, size: 26, color: "2E75B6" })],
                    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 1 } },
                    spacing: { before: 240, after: 120 }
                  }),
                  new Paragraph({
                    children: [new TextRun({ text: "SQL Developer", bold: true, size: 18 })],
                    spacing: { after: 30 }
                  }),
                  new Paragraph({
                    children: [new TextRun({ text: "Megalab", size: 18, color: "666666" })],
                    spacing: { after: 60 }
                  }),
                  new Paragraph({
                    children: [new TextRun({ text: "PL/SQL Developer", bold: true, size: 18 })],
                    spacing: { after: 30 }
                  }),
                  new Paragraph({
                    children: [new TextRun({ text: "Megalab", size: 18, color: "666666" })],
                    spacing: { after: 120 }
                  }),
                  new Paragraph({
                    children: [new TextRun({ text: "React Developer", bold: true, size: 18 })],
                    spacing: { after: 30 }
                  }),
                  new Paragraph({
                    children: [new TextRun({ text: "IT Inkubator", size: 18, color: "666666" })],
                    spacing: { after: 120 }
                  }),

                  // LANGUAGES
                  new Paragraph({
                    children: [new TextRun({ text: "Languages", bold: true, size: 26, color: "2E75B6" })],
                    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 1 } },
                    spacing: { before: 240, after: 120 }
                  }),
                  new Paragraph({
                    children: [
                      new TextRun({ text: "Russian", bold: true, size: 20 }),
                      new TextRun({ text: " – Native", size: 20, color: "555555" })
                    ],
                    spacing: { after: 60 }
                  }),
                  new Paragraph({
                    children: [
                      new TextRun({ text: "English", bold: true, size: 20 }),
                      new TextRun({ text: " – Intermediate", size: 20, color: "555555" })
                    ],
                    spacing: { after: 60 }
                  })
                ]
              }),
              // RIGHT: Experience + Academic Projects
              new TableCell({
                borders: noBorders,
                width: { size: 7080, type: WidthType.DXA },
                margins: { top: 0, bottom: 0, left: 240, right: 0 },
                children: [
                  // EXPERIENCE
                  new Paragraph({
                    children: [new TextRun({ text: "Experience", bold: true, size: 26, color: "2E75B6" })],
                    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 1 } },
                    spacing: { after: 120 }
                  }),

                  jobTitle("Oracle PL/SQL Developer", "DemirBank"),
                  jobMeta("September 2025 | Bishkek, Kyrgyzstan"),
                  new Paragraph({
                    children: [new TextRun({ text: "Create orders for NBKR banking system with Oracle SQL.", size: 20 })],
                    spacing: { after: 60 }
                  }),
                  new Paragraph({
                    children: [new TextRun({ text: "Deployed packages and forms with Oracle PL/SQL", size: 20 })],
                    spacing: { after: 120 }
                  }),

                  // ACADEMIC PROJECTS
                  new Paragraph({
                    children: [new TextRun({ text: "Academic Projects", bold: true, size: 26, color: "2E75B6" })],
                    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 1 } },
                    spacing: { after: 120 }
                  }),

                  new Paragraph({
                    children: [new TextRun({ text: "Portfolio", bold: true, size: 22 })],
                    spacing: { after: 40 }
                  }),
                  jobMeta("2024 – 2025"),
                  new Paragraph({
                    children: [new TextRun({ text: "Developed a portfolio with ReactJS, TailwindCSS, and Firebase.", size: 20 })],
                    spacing: { after: 120 }
                  }),

                  new Paragraph({
                    children: [new TextRun({ text: "Todolist Fullstack", bold: true, size: 22 })],
                    spacing: { after: 40 }
                  }),
                  jobMeta("2023 – 2024"),
                  new Paragraph({
                    children: [new TextRun({ text: "Todolist with CRUD operations and ReactJS", size: 20 })],
                    spacing: { after: 120 }
                  }),

                  new Paragraph({
                    children: [new TextRun({ text: "Explorer of Country", bold: true, size: 22 })],
                    spacing: { after: 40 }
                  }),
                  jobMeta("2022 – 2023"),
                  new Paragraph({
                    children: [new TextRun({ text: "Create explorer of flag and country with restful API and zustand", size: 20 })],
                    spacing: { after: 120 }
                  }),
                ]
              })
            ]
          })
        ]
      })
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(path.join(__dirname, 'Alexey_Folomeev_Resume.docx'), buffer);
  console.log('Done!');
}).catch(console.error);
