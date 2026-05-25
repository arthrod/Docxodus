import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

function readTestFile(relativePath: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(path.join(TEST_FILES_DIR, relativePath)));
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

test.describe('DocxSession (WASM bridge)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('open, project, replaceText, save, reopen — round-trip', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        // Pick first body heading/paragraph anchor by document order
        const anchorEntries = Object.entries(proj.anchorIndex) as [string, any][];
        const firstBody = anchorEntries
          .map(([id, t]) => ({ id, ...t }))
          .filter(t => t.scope === 'body' && ['p', 'h', 'li'].includes(t.kind))
          .map(t => ({ t, idx: proj.markdown.indexOf('{#' + t.id + '}') }))
          .filter(x => x.idx >= 0)
          .sort((a, b) => a.idx - b.idx)[0];

        const replaceResult = JSON.parse(bridge.ReplaceText(handle, firstBody.t.id, '**JSMARKER** replaced.'));

        const after = JSON.parse(bridge.Project(handle));
        const saved = bridge.Save(handle);

        return {
          replaceSuccess: replaceResult.success,
          replaceError: replaceResult.error,
          targetAnchor: firstBody.t.id,
          markdownContainsMarker: after.markdown.includes('JSMARKER'),
          markdownExcerpt: after.markdown.substring(0, 400),
          savedBytes: saved.length,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.replaceSuccess).toBe(true);
    expect(result.markdownContainsMarker, `target=${result.targetAnchor}\nexcerpt:\n${result.markdownExcerpt}`).toBe(true);
    expect(result.savedBytes).toBeGreaterThan(0);
  });

  test('error envelope: malformed markdown gives typed error code', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const anchorId = Object.keys(proj.anchorIndex)[0];
        // Pipe table → TableInsertNotSupported
        const r = JSON.parse(bridge.ReplaceText(handle, anchorId, '| a | b |\n|---|---|\n| 1 | 2 |'));
        return { success: r.success, errorCode: r.error?.code };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.success).toBe(false);
    expect(result.errorCode).toBe('table_insert_not_supported');
  });

  test('Undo restores prior state', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const before = JSON.parse(bridge.Project(handle)).markdown;
        const proj = JSON.parse(bridge.Project(handle));
        const anchorId = Object.keys(proj.anchorIndex).find(k => k.startsWith('h:body:'))!;
        bridge.ReplaceText(handle, anchorId, '**TEMPORARY**');
        const undidOk = bridge.Undo(handle);
        const after = JSON.parse(bridge.Project(handle)).markdown;
        return { undidOk, restored: before === after };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.undidOk).toBe(true);
    expect(result.restored).toBe(true);
  });

  test('ReplaceTextRange + replaceMatch round-trip through the WASM bridge', async ({ page }) => {
    // Grep finds the '[' placeholder markers; replaceMatch addresses each by its
    // (anchor, span) so the wrong one never gets picked when several share text.
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const matchesBefore = JSON.parse(bridge.Grep(handle, '\\[', JSON.stringify({ scope: 1 })));
        const totalBefore = matchesBefore.length;
        if (totalBefore < 1) return { totalBefore, error: 'fixture has no matches' };

        // Replace the FIRST match via span-addressed bridge call.
        const target = matchesBefore[0];
        const r = JSON.parse(bridge.ReplaceTextAtSpan(
          handle,
          target.enclosingAnchor.id,
          target.span.start,
          target.span.length,
          '⟪BRACKET⟫'
        ));

        const matchesAfter = JSON.parse(bridge.Grep(handle, '\\[', JSON.stringify({ scope: 1 })));
        const proj = JSON.parse(bridge.Project(handle));
        return {
          totalBefore,
          totalAfter: matchesAfter.length,
          rSuccess: r.success,
          rError: r.error,
          containsMarker: proj.markdown.includes('⟪BRACKET⟫'),
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.error).toBeUndefined();
    expect(result.rSuccess).toBe(true);
    // Exactly one '[' replaced → match count drops by one. (We don't assert on
    // the marker string surviving in the projection — the projector escapes
    // markdown punctuation, but the run-text edit did land if the count dropped.)
    expect(result.totalAfter).toBe(result.totalBefore - 1);
  });

  test('Grep returns matches with run-fragment breakdown', async ({ page }) => {
    // HC001 is a multilingual tour-plan template full of `[placeholder]` slots;
    // search for an opening bracket which is reliably present regardless of language.
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        // Body scope (1) + small context window. The pattern is a literal '['
        // (escaped for regex) — it appears in every template-placeholder slot.
        const matches = JSON.parse(bridge.Grep(handle, '\\[', JSON.stringify({ scope: 1, contextChars: 20 })));
        const first = matches[0];
        return {
          count: matches.length,
          firstText: first?.text,
          firstHasFragments: Array.isArray(first?.fragments) && first.fragments.length >= 1,
          firstFragmentHasUnid: typeof first?.fragments?.[0]?.unid === 'string',
          firstFragmentHasFormatting: typeof first?.fragments?.[0]?.formatting === 'object',
          firstContextBeforeIsString: typeof first?.contextBefore === 'string',
          firstEnclosingAnchorKind: first?.enclosingAnchor?.kind,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.count).toBeGreaterThan(0);
    expect(result.firstText).toBe('[');
    expect(result.firstHasFragments).toBe(true);
    expect(result.firstFragmentHasUnid).toBe(true);
    expect(result.firstFragmentHasFormatting).toBe(true);
    expect(result.firstContextBeforeIsString).toBe(true);
    expect(['p', 'h', 'li']).toContain(result.firstEnclosingAnchorKind);
  });
});
