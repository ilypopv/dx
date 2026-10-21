// The pdf.js viewer components read the core library from globalThis.pdfjsLib when they are evaluated,
// so this module is imported before them.
import * as lib from 'pdfjs-dist';

(globalThis as any).pdfjsLib = lib;

export const pdfjsLib: any = lib;
