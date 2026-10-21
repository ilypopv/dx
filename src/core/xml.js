'use strict';

// A small XML tree parser for the OOXML parts. It keeps everything it does not understand
// (declaration, comments, processing instructions, the raw attribute text) so that a parsed part
// can be written back byte-for-byte except for the nodes that were edited.
//
// Node shapes:
//   element: { name, attrs: { [name]: value }, children: Node[], raw?: string, selfClosing?: boolean }
//   text:    { text: string }   (decoded)
//   other:   { other: string }  (verbatim: <?xml ...?>, <!-- -->, <![CDATA[ ]]>, <!DOCTYPE>)

const ENTITY = /&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/g;

function decode(text) {
    if (text.indexOf('&') < 0) {
        return text;
    }
    return text.replace(ENTITY, (match, group) => {
        switch (group) {
            case 'amp':
                return '&';
            case 'lt':
                return '<';
            case 'gt':
                return '>';
            case 'quot':
                return '"';
            case 'apos':
                return "'";
        }
        const code = group[1] === 'x' ? parseInt(group.slice(2), 16) : parseInt(group.slice(1), 10);
        return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : match;
    });
}

function encodeText(text) {
    return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function encodeAttr(text) {
    return encodeText(text).replace(/"/g, '&quot;');
}

const ATTR = /([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g;

function parseAttrs(source) {
    const attrs = {};
    ATTR.lastIndex = 0;
    let m;
    while ((m = ATTR.exec(source))) {
        attrs[m[1]] = decode(m[3] !== undefined ? m[3] : m[4]);
    }
    return attrs;
}

function parseXml(source) {
    const root = { name: '#document', attrs: {}, children: [] };
    const stack = [root];
    let i = 0;
    const n = source.length;
    while (i < n) {
        const lt = source.indexOf('<', i);
        if (lt < 0) {
            stack[stack.length - 1].children.push({ text: decode(source.slice(i)) });
            break;
        }
        if (lt > i) {
            stack[stack.length - 1].children.push({ text: decode(source.slice(i, lt)) });
        }
        const next = source[lt + 1];
        if (next === '?' || next === '!') {
            let end;
            if (source.startsWith('<!--', lt)) {
                end = source.indexOf('-->', lt + 4) + 3;
            } else if (source.startsWith('<![CDATA[', lt)) {
                end = source.indexOf(']]>', lt + 9) + 3;
            } else {
                end = source.indexOf('>', lt + 2) + 1;
            }
            if (end <= lt) {
                throw new Error('Unterminated XML markup.');
            }
            const raw = source.slice(lt, end);
            if (raw.startsWith('<![CDATA[')) {
                stack[stack.length - 1].children.push({ text: raw.slice(9, -3), cdata: true });
            } else {
                stack[stack.length - 1].children.push({ other: raw });
            }
            i = end;
            continue;
        }
        // Find the end of the tag, skipping quoted attribute values.
        let j = lt + 1;
        let quote = 0;
        while (j < n) {
            const c = source.charCodeAt(j);
            if (quote) {
                if (c === quote) {
                    quote = 0;
                }
            } else if (c === 34 || c === 39) {
                quote = c;
            } else if (c === 62) {
                break;
            }
            j++;
        }
        if (j >= n) {
            throw new Error('Unterminated XML tag.');
        }
        const body = source.slice(lt + 1, j);
        i = j + 1;
        if (body[0] === '/') {
            const name = body.slice(1).trim();
            while (stack.length > 1) {
                const open = stack.pop();
                if (open.name === name) {
                    break;
                }
            }
            continue;
        }
        const selfClosing = body.endsWith('/');
        const inner = selfClosing ? body.slice(0, -1) : body;
        const space = inner.search(/\s/);
        const name = space < 0 ? inner : inner.slice(0, space);
        const rawAttrs = space < 0 ? '' : inner.slice(space);
        const element = { name, attrs: parseAttrs(rawAttrs), children: [], raw: rawAttrs, selfClosing };
        stack[stack.length - 1].children.push(element);
        if (!selfClosing) {
            stack.push(element);
        }
    }
    return root;
}

function attrsText(node) {
    // Untouched elements keep their original attribute text, edited ones are re-encoded.
    if (node.raw !== undefined && !node.dirtyAttrs) {
        return node.raw;
    }
    let out = '';
    for (const key of Object.keys(node.attrs)) {
        out += ` ${key}="${encodeAttr(String(node.attrs[key]))}"`;
    }
    return out;
}

function serialize(node) {
    const parts = [];
    const walk = current => {
        if (current.other !== undefined) {
            parts.push(current.other);
        } else if (current.text !== undefined) {
            parts.push(current.cdata ? `<![CDATA[${current.text}]]>` : encodeText(current.text));
        } else if (current.name === '#document') {
            current.children.forEach(walk);
        } else if (!current.children.length && current.selfClosing !== false) {
            parts.push(`<${current.name}${attrsText(current)}/>`);
        } else {
            parts.push(`<${current.name}${attrsText(current)}>`);
            current.children.forEach(walk);
            parts.push(`</${current.name}>`);
        }
    };
    walk(node);
    return parts.join('');
}

function element(name, attrs, children) {
    return { name, attrs: attrs || {}, children: children || [], dirtyAttrs: true };
}

function setAttr(node, name, value) {
    node.attrs[name] = value;
    node.dirtyAttrs = true;
}

// Child helpers. Names are matched with their namespace prefix (w:p, a:blip, ...).
function child(node, name) {
    if (!node) {
        return undefined;
    }
    for (const c of node.children) {
        if (c.name === name) {
            return c;
        }
    }
    return undefined;
}

function children(node, name) {
    return node ? node.children.filter(c => c.name === name) : [];
}

// The first descendant with the name, depth first.
function find(node, name) {
    if (!node || !node.children) {
        return undefined;
    }
    for (const c of node.children) {
        if (c.name === name) {
            return c;
        }
        const deep = find(c, name);
        if (deep) {
            return deep;
        }
    }
    return undefined;
}

function textOf(node) {
    if (!node) {
        return '';
    }
    if (node.text !== undefined) {
        return node.text;
    }
    return node.children ? node.children.map(textOf).join('') : '';
}

module.exports = { parseXml, serialize, element, setAttr, child, children, find, textOf, decode };
