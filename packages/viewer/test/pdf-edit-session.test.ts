import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { buildPdf } from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";

/*
 * The main-thread geometry cache of ACTION-825, task 36: elementsAtSync
 * answers hover from the last elements read per page, without waiting
 * behind the session's queue, and the cache follows committed changes.
 */

async function settle(check: () => boolean, label: string): Promise<void> {
  for (let attempt = 0; attempt < 200; attempt += 1) {
    if (check()) return;
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
  assert.fail(`timed out waiting for ${label}`);
}

describe("geometry cache (overlay primitives)", () => {
  it("answers from the cache without awaiting and reports the cached revision", async () => {
    const { session, end } = await pdfSession(
      await buildPdf([
        { texts: [{ text: "Hover me", x: 72, y: 700, fontSize: 24 }] },
        { text: "Second" },
      ]),
    );
    try {
      // Nothing read yet: no items, the current revision, no cached pages.
      const cold = session.elementsAtSync(0, { x: 80, y: 85 });
      assert.deepEqual(cold.items, []);
      assert.equal(cold.revision, 0);
      assert.deepEqual(session.cachedPages, []);

      const page = await session.getElements({ pageIndex: 0 });
      assert.deepEqual(session.cachedPages, [0]);
      const centre = {
        x: page.items[0]!.bounds.x + page.items[0]!.bounds.width / 2,
        y: page.items[0]!.bounds.y + page.items[0]!.bounds.height / 2,
      };
      const sync = session.elementsAtSync(0, centre);
      const async = await session.elementsAt(0, centre);
      assert.deepEqual(sync.items, async.items);
      assert.equal(sync.sessionId, session.sessionId);
      assert.equal(sync.revision, 0);
      assert.deepEqual(session.elementsAtSync(0, { x: 1, y: 1 }).items, []);
      // Narrower queries do not feed the cache.
      await session.getElements({ pageIndex: 1, kinds: ["text"] });
      assert.deepEqual(session.cachedPages, [0]);
    } finally {
      await end();
    }
  });

  it("keeps answering while a change is queued and refreshes afterwards", async () => {
    const { session, end } = await pdfSession(
      await buildPdf([
        { texts: [{ text: "Old", x: 72, y: 700, fontSize: 24 }] },
      ]),
    );
    try {
      await session.getElements({ pageIndex: 0 });
      const pending = session.insertTextBox({
        pageIndex: 0,
        rect: { x: 300, y: 300, width: 200, height: 40 },
        text: "New box",
      });
      // The apply is queued: hover still answers from revision 0.
      const during = session.elementsAtSync(0, { x: 320, y: 310 });
      assert.equal(during.revision, 0);
      assert.deepEqual(during.items, []);
      const receipt = await pending;
      assert.equal(receipt.revision, 1);
      // The changed page is dropped, then read again in the background.
      await settle(
        () =>
          session.cachedPages.includes(0) &&
          session.elementsAtSync(0, { x: 320, y: 310 }).revision === 1,
        "the cache to refresh",
      );
      const after = session.elementsAtSync(0, { x: 320, y: 310 });
      assert.equal(after.items[0]?.id, receipt.createdIds[0]);
      assert.equal(after.items[0]?.kind, "textBox");
      // A dry run changes nothing and keeps the cache.
      await session.insertTextBox(
        {
          pageIndex: 0,
          rect: { x: 10, y: 10, width: 50, height: 20 },
          text: "dry",
        },
        { dryRun: true },
      );
      assert.deepEqual(session.cachedPages, [0]);
      assert.equal(session.elementsAtSync(0, { x: 320, y: 310 }).revision, 1);
      // A deleted page leaves the cache.
      await session.insertPage({ index: 1 });
      await session.getElements({ pageIndex: 1 });
      assert.deepEqual(session.cachedPages, [0, 1]);
      await session.deletePage({ pageIndex: 1 });
      await settle(() => !session.cachedPages.includes(1), "page 1 to drop");
    } finally {
      await end();
    }
  });
});
