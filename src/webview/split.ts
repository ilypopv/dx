// Splits a paragraph between two of its lines, the way Word breaks a page inside a paragraph.
//
// The first part keeps the lines that end above `limit` (a screen y). Word does not leave a single line of a
// paragraph alone at the bottom of a page, nor carry a single line over to the next one: such a break moves one line
// more, or the whole paragraph. Pieces of text and styled spans that the cut goes through are copied to both parts
// with their slot (data-s), so the parts can be put back together when the paragraph is saved.

interface Chars {
    count: number;
    rect(index: number): DOMRect;
    // The text node and offset a character index stands for.
    at(index: number): [Text, number];
}

function chars(p: HTMLElement): Chars {
    const nodes: Text[] = [];
    const starts: number[] = [];
    let count = 0;
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        const t = n as Text;
        if (!t.data.length || (t.parentElement && t.parentElement.closest('.dx-marker'))) {
            continue;
        }
        nodes.push(t);
        starts.push(count);
        count += t.data.length;
    }
    const range = document.createRange();
    const at = (index: number): [Text, number] => {
        let lo = 0;
        let hi = nodes.length - 1;
        while (lo < hi) {
            const mid = (lo + hi + 1) >> 1;
            if (starts[mid] <= index) {
                lo = mid;
            } else {
                hi = mid - 1;
            }
        }
        return [nodes[lo], index - starts[lo]];
    };
    const rect = (index: number): DOMRect => {
        // A collapsed character has no box: the next one that has is on the same line.
        for (let i = index; i < Math.min(count, index + 40); i++) {
            const [node, offset] = at(i);
            range.setStart(node, offset);
            range.setEnd(node, offset + 1);
            const r = range.getBoundingClientRect();
            if (r.height > 0) {
                return r;
            }
        }
        const [node, offset] = at(Math.min(index, count - 1));
        range.setStart(node, offset);
        range.setEnd(node, offset + 1);
        return range.getBoundingClientRect();
    };
    return { count, rect, at };
}

// The first index in [from, to) for which `test` holds, assuming it holds from some index on; `to` if never.
function first(from: number, to: number, test: (i: number) => boolean): number {
    let lo = from;
    let hi = to;
    while (lo < hi) {
        const mid = (lo + hi) >> 1;
        if (test(mid)) {
            hi = mid;
        } else {
            lo = mid + 1;
        }
    }
    return lo;
}

// Returns the paragraph that now holds the lines after the cut, or null when the paragraph is not to be split
// (it fits, nothing of it fits, or a split would leave a single line on either page).
export function splitParagraph(p: HTMLElement, limit: number, scale: number): HTMLElement | null {
    const c = chars(p);
    if (c.count < 2) {
        return null;
    }
    // Lines are as tall as their line height, which is more than the box of their text.
    const lineHeight = parseFloat(getComputedStyle(p).lineHeight);
    const bottomOf = (i: number) => {
        const r = c.rect(i);
        const lead = isFinite(lineHeight) ? Math.max(0, (lineHeight * scale - r.height) / 2) : 0;
        return r.bottom + lead;
    };
    const sameLine = 2 * scale;

    const over = first(0, c.count, i => bottomOf(i) > limit);
    if (over === 0 || over >= c.count) {
        return null;
    }
    const lineStart = (i: number) => {
        const top = c.rect(i).top;
        return first(0, i, j => c.rect(j).top > top - sameLine);
    };
    let cut = lineStart(over);
    const secondLine = first(0, c.count, i => c.rect(i).top > c.rect(0).top + sameLine);
    const lastLine = first(0, c.count, i => c.rect(i).top > c.rect(c.count - 1).top - sameLine);
    if (secondLine >= c.count) {
        return null;
    }
    if (cut >= lastLine) {
        // Only the last line would go on: take the line before it along.
        cut = lineStart(cut - 1);
    }
    if (cut <= secondLine) {
        // A single line would stay behind.
        return null;
    }

    const [node, offset] = c.at(cut);
    const range = document.createRange();
    range.setStart(node, offset);
    range.setEnd(p, p.childNodes.length);
    const tail = range.extractContents();
    const rest = p.cloneNode(false) as HTMLElement;
    rest.classList.add('dx-cont');
    rest.style.textIndent = '0pt';
    rest.style.marginTop = '0';
    rest.appendChild(tail);
    // Space after belongs to the end of the paragraph, the keep-with-next rule to its last line.
    p.style.marginBottom = '0';
    p.classList.remove('dx-keep');
    if (p.style.textAlign === 'justify') {
        // The line before a page break is a line in the middle of the paragraph: it is justified, too.
        p.style.textAlignLast = 'justify';
    }
    p.after(rest);
    return rest;
}
