// Cuts a sheet that holds more than one page of content into pages.
//
// Word records where it started a page (lastRenderedPageBreak), but not every break, so a sheet between two records
// can hold several pages. How many is what its content measures against the page area (to the nearest page: our text
// is a few percent off Word's). The pages are then filled the way Word fills them, one after the other up to the
// bottom margin: a break falls between two lines of a paragraph, between two rows of a table, or before the block
// that does not fit. A heading that asks to stay with the next block is not left alone at the bottom of a page.
// Measuring is done on the laid out DOM, in whatever zoom it is shown.

import { splitParagraph } from './split';

export interface Measure {
    // Height of the area of the sheet the blocks flow in, as measured on screen.
    capacity: number;
    // Screen pixels per CSS pixel of the sheet.
    scale: number;
}

export interface PageHost {
    // undefined when the sheet is not laid out (hidden).
    measure(sheet: HTMLElement): Measure | undefined;
    bodyOf(sheet: HTMLElement): HTMLElement;
    // A new empty sheet right after `sheet`, of the same section.
    follow(sheet: HTMLElement): HTMLElement;
    // Cuts a table between rows at `limit` (screen y); false when no row can start a new part.
    splitTable(wrap: HTMLElement, limit: number): boolean;
    label(sheet: HTMLElement): number;
}

// While a sheet is still being filled it is only cut once it clearly holds this many pages.
const OPEN_MIN_PAGES = 3;

// Returns the sheet that stays open for more blocks (the last piece).
export function cutSheet(host: PageHost, sheet: HTMLElement, final: boolean): HTMLElement {
    const first = host.measure(sheet);
    const last = host.bodyOf(sheet).lastElementChild;
    if (!first || !last) {
        return sheet;
    }
    const ratio =
        (last.getBoundingClientRect().bottom - host.bodyOf(sheet).getBoundingClientRect().top) / first.capacity;
    const pages = Math.round(ratio);
    if (final ? pages < 2 : ratio < OPEN_MIN_PAGES) {
        return sheet;
    }
    console.log('[dx] sheet ' + host.label(sheet) + ' holds ' + ratio.toFixed(2) + ' pages, cut into ' + pages);

    let cur = sheet;
    for (let left = pages; left > 1; left--) {
        const measure = host.measure(cur);
        const body = host.bodyOf(cur);
        if (!measure) {
            break;
        }
        // Everything below this line does not fit on the page.
        const limit = body.getBoundingClientRect().top + measure.capacity;

        let kids = Array.from(body.children) as HTMLElement[];
        const table = kids.find(c => {
            if (!c.classList.contains('dx-table-wrap')) {
                return false;
            }
            const r = c.getBoundingClientRect();
            return r.top < limit - 1 && r.bottom > limit + 1;
        });
        if (table && host.splitTable(table, limit)) {
            kids = Array.from(body.children) as HTMLElement[];
        }

        let at = kids.findIndex(c => c.getBoundingClientRect().bottom > limit + 0.5);
        if (at < 0) {
            break;
        }
        if (kids[at].classList.contains('dx-p') && splitParagraph(kids[at], limit, measure.scale)) {
            at++;
            kids = Array.from(body.children) as HTMLElement[];
        }
        // A page keeps at least one block.
        at = Math.max(at, 1);
        // A heading stays with the block that follows it.
        while (at > 1 && kids[at - 1].classList.contains('dx-keep')) {
            at--;
        }
        if (at >= kids.length) {
            break;
        }
        const next = host.follow(cur);
        host.bodyOf(next).append(...kids.slice(at));
        cur = next;
    }
    return cur;
}
