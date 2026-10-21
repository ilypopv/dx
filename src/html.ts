// JSON for an inline <script>: no sequence in the output can close the script or break the JS string/line grammar.
export function safeJson(value: any): string {
    return JSON.stringify(value)
        .replace(/</g, '\\u003c')
        .replace(/>/g, '\\u003e')
        .replace(/&/g, '\\u0026')
        .replace(/\u2028/g, '\\u2028')
        .replace(/\u2029/g, '\\u2029');
}

export function makeNonce(): string {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    const bytes = new Uint8Array(32);
    (globalThis as any).crypto.getRandomValues(bytes);
    let out = '';
    for (let i = 0; i < bytes.length; i++) {
        out += chars[bytes[i] % chars.length];
    }
    return out;
}

// The xl policy plus what the document viewers need: the pdf.js worker is a same-origin blob worker
// (worker-src blob:), the images of a document are blob URLs, and the pdf.js cmaps are fetched from the extension.
export function csp(cspSource: string, nonce: string): string {
    return `default-src 'none'; style-src ${cspSource} 'unsafe-inline'; script-src ${cspSource} 'nonce-${nonce}'; img-src ${cspSource} data: blob:; font-src ${cspSource} data: blob:; worker-src blob:; connect-src ${cspSource};`;
}
