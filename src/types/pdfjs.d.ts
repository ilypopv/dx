// The pdf.js worker is inlined into the webview bundle as a string (webpack asset/source)
// and turned into a same-origin blob worker at runtime, so no cross-origin request is ever made.
declare module 'pdfjs-dist/build/pdf.worker.min.mjs' {
    const source: string;
    export default source;
}

declare module 'pdfjs-dist/web/pdf_viewer.mjs' {
    export const EventBus: any;
    export const PDFLinkService: any;
    export const PDFFindController: any;
    export const PDFViewer: any;
    export const FindState: any;
    export const LinkTarget: any;
}

declare module 'pdfjs-dist' {
    const lib: any;
    export = lib;
}

declare module '*.js';
