import assert from "node:assert/strict";
import { describe, it } from "node:test";

import {
  buildPdf,
  fixturePdfium,
  type FixtureText,
} from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";

const wrappedLines = [
  "Since 2013 our independent testing has tracked",
  "software quality across teams and organisations",
  "and shared the results with practitioners.",
];

function rows(lines: readonly string[]): FixtureText[] {
  return lines.map((text, index) => ({
    text,
    x: 72,
    y: 700 - index * 20,
    fontSize: 11,
  }));
}

describe("conservative imported PDF paragraph recognition", () => {
  for (const [position, neighbor] of [
    [
      "distant same baseline",
      { text: "- Joel and Lalit", x: 470, y: 680, fontSize: 11 },
    ],
    [
      "nearby different baseline",
      { text: "- Joel and Lalit", x: 320, y: 675, fontSize: 11 },
    ],
    [
      "large numeral beyond the body text neighborhood",
      { text: "01", x: 350, y: 700, fontSize: 54 },
    ],
  ] as const) {
    it(`keeps a body paragraph beside ${position}`, async () => {
      const { session, end } = await pdfSession(
        await buildPdf([
          {
            texts: [...rows(wrappedLines), neighbor],
          },
        ]),
      );
      try {
        const elements = (await session.getElements({ pageIndex: 0 })).items;
        const first = elements[0];
        assert.ok(first);
        const paragraph = (await session.getTextParagraph(first.id)).item;
        assert.ok(paragraph);
        assert.equal(paragraph.text, wrappedLines.join(" "));
        assert.equal(paragraph.memberIds.length, 3);
        assert.equal(elements.at(-1)?.textEditingTarget, undefined);
      } finally {
        await end();
      }
    });
  }

  for (const [placement, sidebarX] of [
    ["outside", 20],
    ["intersecting", 100],
  ] as const) {
    it(`handles a rotated sidebar ${placement} the body paragraph without absorbing it`, async () => {
      const pdfium = await fixturePdfium();
      const original = await buildPdf([
        {
          texts: [
            ...rows(wrappedLines),
            { text: "Rotated sidebar", x: 0, y: 0, fontSize: 11 },
          ],
        },
      ]);
      const document = pdfium.openDocument(original);
      const page = pdfium.lib.FPDF_LoadPage(document.handle, 0);
      let rotated: Uint8Array;
      try {
        const object = pdfium.lib.FPDFPage_GetObject(page, 3);
        pdfium.lib.FPDFPageObj_Transform(object, 0, 1, -1, 0, sidebarX, 650);
        pdfium.lib.FPDFPage_GenerateContent(page);
        rotated = document.save("full");
      } finally {
        pdfium.lib.FPDF_ClosePage(page);
        document.close();
      }
      const { session, end } = await pdfSession(rotated);
      try {
        const elements = (await session.getElements({ pageIndex: 0 })).items;
        const first = elements[0];
        const sidebar = elements[3];
        assert.ok(first && sidebar);
        assert.ok(sidebar.rotation);
        const paragraph = (await session.getTextParagraph(first.id)).item;
        if (placement === "intersecting") {
          assert.equal(paragraph, undefined);
          assert.ok(
            elements.every(
              (element) => element.textEditingTarget === undefined,
            ),
          );
          return;
        }
        assert.ok(paragraph);
        assert.equal(paragraph.text, wrappedLines.join(" "));
        await session.replaceParagraphText({
          target: paragraph.id,
          text: "Body edited safely.",
        });
        assert.deepEqual((await session.getElement(sidebar.id)).item, sidebar);
      } finally {
        await end();
      }
    });
  }

  it("compares each continuation's alignment with the first row, not pairwise", async () => {
    const original = await buildPdf([
      {
        texts: rows(wrappedLines).map((row, index) => ({
          ...row,
          x: index === 0 ? 72 : index === 1 ? 71 : 73,
          // Native ink overlaps vertically at this valid compact leading.
          y: 700 - index * 9,
        })),
      },
    ]);
    const { session, end } = await pdfSession(original);
    try {
      const first = (await session.getElements({ pageIndex: 0 })).items[0];
      assert.ok(first);
      const paragraph = (await session.getTextParagraph(first.id)).item;
      assert.ok(paragraph);
      assert.equal(paragraph.text, wrappedLines.join(" "));
      assert.equal(paragraph.memberIds.length, 3);
    } finally {
      await end();
    }
  });

  const refused: readonly {
    name: string;
    texts: readonly FixtureText[];
  }[] = [
    {
      name: "adjacent table cells sharing row baselines",
      texts: [
        ...rows(wrappedLines),
        ...rows(["Value one", "Value two", "Value three"]).map((row) => ({
          ...row,
          x: 320,
        })),
      ],
    },
    {
      name: "list text with separate bullet objects",
      texts: [
        ...rows(wrappedLines),
        ...rows(["•", "•", "•"]).map((row) => ({ ...row, x: 54 })),
      ],
    },
    {
      name: "short paragraph endings followed immediately by another paragraph",
      texts: rows(["A short paragraph ends.", ...wrappedLines]),
    },
    {
      name: "a first-line indent followed by aligned continuation rows",
      texts: rows(wrappedLines).map((row, index) => ({
        ...row,
        x: index === 0 ? 90 : 72,
      })),
    },
    {
      name: "matching fragments on either side of an excluded styled middle row",
      texts: rows([
        wrappedLines[0]!,
        wrappedLines[1]!,
        "An emphasized middle line belongs to this section",
        wrappedLines[0]!,
        wrappedLines[1]!,
      ]).map((row, index) =>
        index === 2 ? { ...row, font: "Helvetica-Bold" } : row,
      ),
    },
    {
      name: "a connected excluded leading-whitespace row",
      texts: rows([` ${wrappedLines[0]}`, wrappedLines[1]!, wrappedLines[2]!]),
    },
    ...["•", "*", "1.", "(1)", "a)", "–"].map((marker) => ({
      name: `separate list items with in-object ${marker} markers`,
      texts: rows([
        `${marker} Check the first stable setting`,
        `${marker} Check the next stable setting`,
        `${marker} Check the last stable setting`,
      ]),
    })),
    {
      name: "equal-width standalone dates",
      texts: rows(["2024-01-01", "2024-02-01", "2024-03-01"]),
    },
  ];

  for (const fixture of refused) {
    it(`keeps ${fixture.name} independently editable`, async () => {
      const original = await buildPdf([{ texts: fixture.texts }]);
      const { session, end } = await pdfSession(original);
      try {
        const elements = (await session.getElements({ pageIndex: 0 })).items;
        assert.equal(elements.length, fixture.texts.length);
        for (const element of elements) {
          assert.equal(
            element.textEditingTarget,
            undefined,
            element.text ?? element.id,
          );
          assert.equal(
            (await session.getTextParagraph(element.id)).item,
            undefined,
            element.text ?? element.id,
          );
          assert.ok(element.operations.includes("replaceText"));
        }
        assert.equal(session.state.dirty, false);
        assert.deepEqual((await session.save()).bytes, original);
      } finally {
        await end();
      }
    });
  }

  it("refuses a line-end hyphen extracted as U+0002 without rewriting its text", async () => {
    const original = await buildPdf([
      {
        texts: rows([
          "The complete paragraph starts with hyphen-",
          "ated words which continue in the next line",
          "and then finish here.",
        ]),
      },
    ]);
    const { session, end } = await pdfSession(original);
    try {
      const elements = (await session.getElements({ pageIndex: 0 })).items;
      assert.ok(elements[0]?.text?.endsWith("\u0002"));
      for (const element of elements) {
        assert.equal(element.textEditingTarget, undefined);
        assert.equal(
          (await session.getTextParagraph(element.id)).item,
          undefined,
        );
      }
      assert.deepEqual((await session.save()).bytes, original);
    } finally {
      await end();
    }
  });

  it("retains complete nearly filled wrapped lines with a shorter final line and normalizes their size", async () => {
    const original = await buildPdf([
      {
        texts: rows(wrappedLines).map((row) => ({
          ...row,
          fontSize: 10.123456,
        })),
      },
    ]);
    const { session, end } = await pdfSession(original);
    try {
      const elements = (await session.getElements({ pageIndex: 0 })).items;
      const first = elements[0];
      assert.ok(first);
      const paragraph = (await session.getTextParagraph(first.id)).item;
      assert.ok(paragraph);
      assert.equal(paragraph.text, wrappedLines.join(" "));
      assert.equal(paragraph.textStyle.fontSize, 10.123);
      assert.deepEqual(
        paragraph.memberIds,
        elements.map((element) => element.id),
      );
      for (const element of elements)
        assert.equal(element.textEditingTarget, paragraph.id);
      assert.deepEqual((await session.save()).bytes, original);
    } finally {
      await end();
    }
  });
});
