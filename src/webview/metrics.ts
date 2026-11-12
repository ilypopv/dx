// Font measuring for the layout engine in the webview, the way Word measures: advance widths without kerning or
// ligatures (a canvas with both off), and the height of a line of a font as ascent + descent + line gap, which is
// the browser's `line-height: normal`.

const GENERIC_SERIF = /times|georgia|cambria|garamond|serif|book|minion|palatino/i;
const GENERIC_MONO = /courier|consolas|mono|menlo/i;

export const DEFAULT_STACK = 'Calibri, Carlito, sans-serif';

export function fontStack(name: string | undefined): string {
    if (!name) {
        return '';
    }
    const generic = GENERIC_MONO.test(name) ? 'monospace' : GENERIC_SERIF.test(name) ? 'serif' : 'sans-serif';
    return `"${name.replace(/["\\]/g, '')}", ${generic}`;
}

export interface Font {
    family: string;
    size: number;
    b?: boolean;
    i?: boolean;
}

export interface Measure {
    width(text: string, font: Font): number;
    metrics(font: Font): { ascent: number; descent: number; natural: number };
}

const SIZE = 100;

function css(font: Font): string {
    return `${font.i ? 'italic ' : ''}${font.b ? 'bold ' : ''}${SIZE}px ${fontStack(font.family) || DEFAULT_STACK}`;
}

const ratios = new Map<string, number>();
let probe: HTMLElement | undefined;

// The height of a line of the font at `line-height: normal`, in ems.
export function lineRatio(stack: string, bold = false, italic = false): number {
    const key = stack + (bold ? 'b' : '') + (italic ? 'i' : '');
    const known = ratios.get(key);
    if (known !== undefined) {
        return known;
    }
    if (!probe) {
        probe = document.createElement('div');
        probe.style.cssText = 'position:absolute;left:-9999px;top:0;visibility:hidden;white-space:nowrap';
        probe.appendChild(document.createElement('span'));
        document.body.appendChild(probe);
    }
    const span = probe.firstChild as HTMLElement;
    span.style.cssText = `display:inline-block;font:${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${SIZE}px ${stack};line-height:normal`;
    span.textContent = 'Hg';
    const ratio = span.getBoundingClientRect().height / SIZE;
    const value = ratio > 0.5 && ratio < 3 ? ratio : 1.15;
    ratios.set(key, value);
    return value;
}

export function createMeasure(): Measure {
    const ctx = document.createElement('canvas').getContext('2d') as any;
    // Word sets text without kerning and without ligatures unless a run asks for them.
    ctx.fontKerning = 'none';
    ctx.textRendering = 'optimizeSpeed';
    ctx.fontVariantLigatures = 'none';
    const widths = new Map<string, number>();
    const lines = new Map<string, { ascent: number; descent: number; natural: number }>();
    let current = '';
    const use = (font: Font) => {
        const f = css(font);
        if (f !== current) {
            ctx.font = f;
            current = f;
        }
        return f;
    };
    return {
        width(text, font) {
            const key = css(font) + '\u0000' + text;
            let w = widths.get(key);
            if (w === undefined) {
                use(font);
                w = ctx.measureText(text).width / SIZE;
                widths.set(key, w);
            }
            return w * font.size;
        },
        metrics(font) {
            const key = css(font);
            let m = lines.get(key);
            if (!m) {
                use(font);
                const box = ctx.measureText('Hg');
                const natural = lineRatio(fontStack(font.family) || DEFAULT_STACK, !!font.b, !!font.i);
                const descent = (box.fontBoundingBoxDescent || 0.216 * SIZE) / SIZE;
                // The line gap sits above the text, in the ascent.
                m = { ascent: natural - descent, descent, natural };
                lines.set(key, m);
            }
            return { ascent: m.ascent * font.size, descent: m.descent * font.size, natural: m.natural * font.size };
        }
    };
}
