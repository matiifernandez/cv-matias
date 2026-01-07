const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  AlignmentType,
  WidthType,
  BorderStyle,
  ExternalHyperlink,
  LevelFormat,
  ShadingType,
} = require("docx");
const fs = require("fs");
require('dotenv').config();

const PRIMARY = "2563EB";
const DARK = "1E293B";
const TEXT = "334155";
const LIGHT_TEXT = "64748B";
const SIDEBAR_BG = "F1F5F9";
const WHITE = "FFFFFF";

const noBorder = { style: BorderStyle.NONE, size: 0, color: WHITE };
const noBorders = {
  top: noBorder,
  bottom: noBorder,
  left: noBorder,
  right: noBorder,
};

function sectionLeft(title) {
  return new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [
      new TextRun({
        text: title,
        bold: true,
        size: 24,
        color: PRIMARY,
        font: "Calibri",
        allCaps: true,
      }),
    ],
  });
}

function sectionRight(title) {
  return new Paragraph({
    spacing: { before: 180, after: 80 },
    children: [
      new TextRun({
        text: title,
        bold: true,
        size: 20,
        color: DARK,
        font: "Calibri",
        allCaps: true,
      }),
    ],
  });
}

function jobEntry(title, company, dates, location, bullets, tag = null) {
  const children = [
    new Paragraph({
      spacing: { before: 120, after: 30 },
      children: [
        new TextRun({
          text: title,
          bold: true,
          size: 22,
          color: DARK,
          font: "Calibri",
        }),
        tag
          ? new TextRun({
              text: `  ${tag}`,
              size: 18,
              color: PRIMARY,
              font: "Calibri",
            })
          : null,
      ].filter(Boolean),
    }),
    new Paragraph({
      spacing: { after: 30 },
      children: [
        new TextRun({
          text: company,
          italics: true,
          size: 20,
          color: TEXT,
          font: "Calibri",
        }),
        new TextRun({
          text: `  •  ${location}`,
          size: 18,
          color: LIGHT_TEXT,
          font: "Calibri",
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 50 },
      children: [
        new TextRun({
          text: dates,
          size: 18,
          color: LIGHT_TEXT,
          font: "Calibri",
        }),
      ],
    }),
  ];

  bullets.forEach((bullet) => {
    children.push(
      new Paragraph({
        numbering: { reference: "modern-bullets", level: 0 },
        spacing: { after: 40 },
        children: bullet,
      })
    );
  });

  return children;
}

function projectEntry(name, role, description, stack, link = null) {
  const stackLine = [
    new TextRun({
      text: stack,
      size: 18,
      color: LIGHT_TEXT,
      font: "Calibri",
      italics: true,
    }),
  ];
  if (link) {
    stackLine.push(
      new TextRun({
        text: "  •  ",
        size: 18,
        color: LIGHT_TEXT,
        font: "Calibri",
      })
    );
    stackLine.push(
      new ExternalHyperlink({
        children: [
          new TextRun({
            text: link.text,
            color: PRIMARY,
            size: 18,
            font: "Calibri",
          }),
        ],
        link: link.url,
      })
    );
  }

  return [
    new Paragraph({
      spacing: { before: 100, after: 20 },
      children: [
        new TextRun({
          text: name,
          bold: true,
          size: 21,
          color: DARK,
          font: "Calibri",
        }),
        new TextRun({
          text: `  —  ${role}`,
          size: 19,
          color: TEXT,
          font: "Calibri",
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 30 },
      children: [
        new TextRun({
          text: description,
          size: 19,
          color: TEXT,
          font: "Calibri",
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 60 },
      children: stackLine,
    }),
  ];
}

function sidebarItem(label, value, isLink = false, url = null) {
  if (isLink) {
    return new Paragraph({
      spacing: { after: 50 },
      children: [
        new ExternalHyperlink({
          children: [
            new TextRun({
              text: value,
              color: PRIMARY,
              size: 20,
              font: "Calibri",
              bold: true,
            }),
          ],
          link: url,
        }),
      ],
    });
  }
  return new Paragraph({
    spacing: { after: 50 },
    children: [
      new TextRun({
        text: label ? `${label}  ` : "",
        size: 18,
        color: LIGHT_TEXT,
        font: "Calibri",
      }),
      new TextRun({ text: value, size: 20, color: DARK, font: "Calibri" }),
    ],
  });
}

function skillItem(category, skills) {
  return new Paragraph({
    spacing: { after: 70 },
    children: [
      new TextRun({
        text: category,
        bold: true,
        size: 19,
        color: DARK,
        font: "Calibri",
      }),
      new TextRun({ text: "\n", size: 10 }),
      new TextRun({ text: skills, size: 18, color: TEXT, font: "Calibri" }),
    ],
  });
}

function eduItem(title, school, year, details) {
  return [
    new Paragraph({
      spacing: { before: 70, after: 15 },
      children: [
        new TextRun({
          text: title,
          bold: true,
          size: 19,
          color: DARK,
          font: "Calibri",
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 15 },
      children: [
        new TextRun({ text: school, size: 17, color: TEXT, font: "Calibri" }),
        new TextRun({
          text: `  •  ${year}`,
          size: 17,
          color: LIGHT_TEXT,
          font: "Calibri",
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 50 },
      children: [
        new TextRun({
          text: details,
          size: 16,
          color: LIGHT_TEXT,
          font: "Calibri",
          italics: true,
        }),
      ],
    }),
  ];
}

const leftContent = [
  sectionLeft("Profile"),
  new Paragraph({
    spacing: { after: 120 },
    children: [
      new TextRun({
        text: "After 10 years solving real-world problems—from managing restaurant P&L to installing critical infrastructure—I pivoted to software engineering to ",
        size: 21,
        color: TEXT,
        font: "Calibri",
      }),
      new TextRun({
        text: "build systems, not just operate them",
        bold: true,
        size: 21,
        color: DARK,
        font: "Calibri",
      }),
      new TextRun({
        text: ". Recently graduated from ",
        size: 21,
        color: TEXT,
        font: "Calibri",
      }),
      new TextRun({
        text: "Le Wagon Tokyo",
        bold: true,
        size: 21,
        color: PRIMARY,
        font: "Calibri",
      }),
      new TextRun({
        text: ", I combine ",
        size: 21,
        color: TEXT,
        font: "Calibri",
      }),
      new TextRun({
        text: "operational leadership",
        bold: true,
        size: 21,
        color: DARK,
        font: "Calibri",
      }),
      new TextRun({ text: ", ", size: 21, color: TEXT, font: "Calibri" }),
      new TextRun({
        text: "zero-error discipline",
        bold: true,
        size: 21,
        color: DARK,
        font: "Calibri",
      }),
      new TextRun({ text: ", and ", size: 21, color: TEXT, font: "Calibri" }),
      new TextRun({
        text: "crisis management",
        bold: true,
        size: 21,
        color: DARK,
        font: "Calibri",
      }),
      new TextRun({
        text: " with Ruby on Rails. ",
        size: 21,
        color: TEXT,
        font: "Calibri",
      }),
      new TextRun({
        text: "Available immediately.",
        bold: true,
        size: 21,
        color: PRIMARY,
        font: "Calibri",
      }),
    ],
  }),

  sectionLeft("Experience"),

  ...jobEntry(
    "Freelance Web Developer",
    "Self-Employed",
    "Dec 2025 – Present",
    "Tokyo",
    [
      [
        new TextRun({
          text: "Full-Stack: ",
          bold: true,
          size: 19,
          color: DARK,
          font: "Calibri",
        }),
        new TextRun({
          text: "Production apps with ",
          size: 19,
          color: TEXT,
          font: "Calibri",
        }),
        new TextRun({
          text: "Rails, PostgreSQL, AI integration",
          bold: true,
          size: 19,
          color: TEXT,
          font: "Calibri",
        }),
      ],
      [
        new TextRun({
          text: "Projects: ",
          bold: true,
          size: 19,
          color: DARK,
          font: "Calibri",
        }),
        new TextRun({
          text: "Language learning platform, AI UI generator—deployed",
          size: 19,
          color: TEXT,
          font: "Calibri",
        }),
      ],
    ]
  ),

  ...jobEntry(
    "Cinema Display Technician",
    "AEON Cinemas",
    "Jun 2025 – Present",
    "Tokyo",
    [
      [
        new TextRun({
          text: "Adaptation: ",
          bold: true,
          size: 19,
          color: DARK,
          font: "Calibri",
        }),
        new TextRun({
          text: "Non-Japanese speaker in local team",
          size: 19,
          color: TEXT,
          font: "Calibri",
        }),
      ],
    ],
    "Part-time"
  ),

  ...jobEntry(
    "Dewatering Technician",
    "Dewatering Solutions",
    "Nov 2024 – Mar 2025",
    "Perth, Australia",
    [
      [
        new TextRun({
          text: "Systems: ",
          bold: true,
          size: 19,
          color: DARK,
          font: "Calibri",
        }),
        new TextRun({
          text: "Complex installations under ",
          size: 19,
          color: TEXT,
          font: "Calibri",
        }),
        new TextRun({
          text: "strict protocols",
          bold: true,
          size: 19,
          color: TEXT,
          font: "Calibri",
        }),
      ],
    ]
  ),

  ...jobEntry(
    "Protocol Technician",
    "The Alfred Hospital",
    "Apr 2022 – Jan 2023",
    "Melbourne, Australia",
    [
      [
        new TextRun({
          text: "Zero-Error: ",
          bold: true,
          size: 19,
          color: DARK,
          font: "Calibri",
        }),
        new TextRun({
          text: "High-risk sanitization with ",
          size: 19,
          color: TEXT,
          font: "Calibri",
        }),
        new TextRun({
          text: "mission-critical reliability",
          bold: true,
          size: 19,
          color: TEXT,
          font: "Calibri",
        }),
      ],
    ]
  ),

  ...jobEntry(
    "Installer & Logistics",
    "Stack It Shelving",
    "(Apr 2020–Mar 2022) (Feb 2023–Nov 2024)",
    "Melbourne, Australia",
    [
      [
        new TextRun({
          text: "Leadership: ",
          bold: true,
          size: 19,
          color: DARK,
          font: "Calibri",
        }),
        new TextRun({
          text: "Large-scale installations, ",
          size: 19,
          color: TEXT,
          font: "Calibri",
        }),
        new TextRun({
          text: "project coordination",
          bold: true,
          size: 19,
          color: TEXT,
          font: "Calibri",
        }),
      ],
    ]
  ),

  ...jobEntry(
    "Operations Manager",
    "Las Mellis SRL",
    "Jan 2013 – Dec 2018",
    "Córdoba, Argentina",
    [
      [
        new TextRun({
          text: "Business: ",
          bold: true,
          size: 19,
          color: DARK,
          font: "Calibri",
        }),
        new TextRun({ text: "Full ", size: 19, color: TEXT, font: "Calibri" }),
        new TextRun({
          text: "P&L responsibility",
          bold: true,
          size: 19,
          color: TEXT,
          font: "Calibri",
        }),
        new TextRun({
          text: ", profit-sharing model",
          size: 19,
          color: TEXT,
          font: "Calibri",
        }),
      ],
    ]
  ),

  sectionLeft("Projects"),
  ...projectEntry(
    "Kizuna Lingua",
    "Full-Stack Developer",
    "Language learning for couples/friends. AI topics, bilingual conversation analysis.",
    "Rails • PostgreSQL • JavaScript • Heroku",
    { text: "kizunalingua.com", url: "https://www.kizunalingua.com/" }
  ),
  ...projectEntry(
    "UI Forge",
    "Full-Stack Developer",
    "AI-powered UI components creator for developers and designers.",
    "Rails • PostgreSQL • RubyLLM (OpenAI, Gemini, Groq) • Heroku",
    {
      text: "Live Demo",
      url: "https://ai-assistant-matiifernandez-132f8eab454e.herokuapp.com/",
    }
  ),
];

const rightContent = [
  new Paragraph({
    spacing: { after: 120 },
    children: [
      new TextRun({
        text: "SPOUSE VISA",
        bold: true,
        size: 20,
        color: DARK,
        font: "Calibri",
      }),
      new TextRun({
        text: " — No Sponsorship Required",
        size: 17,
        color: TEXT,
        font: "Calibri",
      }),
    ],
  }),

  sectionRight("Contact"),
  sidebarItem("", process.env.MY_EMAIL),
  sidebarItem("", process.env.MY_PHONE),
  sidebarItem("", "Tokyo, Japan"),
  sidebarItem(
    "",
    "LinkedIn",
    true,
    "https://linkedin.com/in/matias-fernandez-jp"
  ),
  sidebarItem("", "GitHub", true, "https://github.com/matiifernandez"),

  sectionRight("Tech Stack"),
  skillItem(
    "Core",
    "Ruby on Rails, PostgreSQL, JavaScript (Stimulus), Tailwind/Bootstrap, HTML/CSS"
  ),
  skillItem(
    "AI & Tools",
    "OpenAI, Groq, Gemini, RubyLLM, Git, Heroku, Docker, Figma"
  ),

  sectionRight("Languages"),
  new Paragraph({
    spacing: { after: 40 },
    children: [
      new TextRun({
        text: "Spanish ",
        bold: true,
        size: 19,
        color: DARK,
        font: "Calibri",
      }),
      new TextRun({ text: "Native", size: 18, color: TEXT, font: "Calibri" }),
    ],
  }),
  new Paragraph({
    spacing: { after: 40 },
    children: [
      new TextRun({
        text: "English ",
        bold: true,
        size: 19,
        color: DARK,
        font: "Calibri",
      }),
      new TextRun({ text: "Business", size: 18, color: TEXT, font: "Calibri" }),
    ],
  }),
  new Paragraph({
    spacing: { after: 70 },
    children: [
      new TextRun({
        text: "Japanese ",
        bold: true,
        size: 19,
        color: DARK,
        font: "Calibri",
      }),
      new TextRun({
        text: "Studying (N4)",
        size: 18,
        color: TEXT,
        font: "Calibri",
      }),
    ],
  }),

  sectionRight("Key Strengths"),
  new Paragraph({
    spacing: { after: 30 },
    children: [
      new TextRun({
        text: "• Operational Leadership",
        size: 18,
        color: TEXT,
        font: "Calibri",
      }),
    ],
  }),
  new Paragraph({
    spacing: { after: 30 },
    children: [
      new TextRun({
        text: "• Zero-Error Mindset",
        size: 18,
        color: TEXT,
        font: "Calibri",
      }),
    ],
  }),
  new Paragraph({
    spacing: { after: 30 },
    children: [
      new TextRun({
        text: "• Crisis Management",
        size: 18,
        color: TEXT,
        font: "Calibri",
      }),
    ],
  }),
  new Paragraph({
    spacing: { after: 30 },
    children: [
      new TextRun({
        text: "• Cross-Cultural Work",
        size: 18,
        color: TEXT,
        font: "Calibri",
      }),
    ],
  }),
  new Paragraph({
    spacing: { after: 70 },
    children: [
      new TextRun({
        text: "• Adaptability",
        size: 18,
        color: TEXT,
        font: "Calibri",
      }),
    ],
  }),

  sectionRight("Education"),
  ...eduItem(
    "Full-Stack Web Dev",
    "Le Wagon Tokyo",
    "2025",
    "MVC, APIs, AI/LLM, Deployment"
  ),
  ...eduItem("Argentina Programa", "UTN Virtual", "2022", "Logic, JavaScript"),
  ...eduItem("Mechatronics", "UTN Argentina", "2013", "Systems, PLCs"),
];

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "modern-bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: {
              paragraph: { indent: { left: 300, hanging: 200 } },
              run: { color: TEXT },
            },
          },
        ],
      },
    ],
  },
  styles: { default: { document: { run: { font: "Calibri", size: 20 } } } },
  sections: [
    {
      properties: {
        page: { margin: { top: 450, right: 0, bottom: 350, left: 600 } },
      },
      children: [
        new Table({
          columnWidths: [10000],
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders: noBorders,
                  width: { size: 10000, type: WidthType.DXA },
                  children: [
                    new Paragraph({
                      spacing: { after: 30 },
                      children: [
                        new TextRun({
                          text: "MATIAS FERNANDEZ",
                          bold: true,
                          size: 50,
                          color: DARK,
                          font: "Calibri",
                        }),
                      ],
                    }),
                    new Paragraph({
                      spacing: { after: 50 },
                      children: [
                        new TextRun({
                          text: "Full-Stack Developer",
                          size: 26,
                          color: PRIMARY,
                          font: "Calibri",
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
        new Table({
          columnWidths: [6200, 3800],
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders: noBorders,
                  width: { size: 6200, type: WidthType.DXA },
                  margins: { right: 250 },
                  children: leftContent,
                }),
                new TableCell({
                  borders: noBorders,
                  width: { size: 3800, type: WidthType.DXA },
                  shading: { fill: SIDEBAR_BG, type: ShadingType.CLEAR },
                  margins: { top: 130, left: 180, right: 130, bottom: 130 },
                  children: rightContent,
                }),
              ],
            }),
          ],
        }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("MATIAS-FERNANDEZ-CV-modern.docx", buffer);
  console.log("CV with Freelance experience created!");
});
