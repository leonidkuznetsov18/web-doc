import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { buildPdf, type FixtureText } from "./fixtures/pdf-builder.js";
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
  const refused: readonly {
    name: string;
    texts: readonly FixtureText[];
  }[] = [
    {
      name: "aligned table cells or columns sharing row baselines",
      texts: [
        ...rows(wrappedLines),
        ...rows(["Value one", "Value two", "Value three"]).map((row) => ({
          ...row,
          x: 350,
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
