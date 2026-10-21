'use strict';

// A reader of the Compound File Binary format (OLE2 / MS-CFB), the container of the Word 97-2003 .doc files.
// Only reading is needed: the directory is resolved into named streams, each stream is returned as a Uint8Array.

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const FREESECT = 0xffffffff;
const ENDOFCHAIN = 0xfffffffe;
const MAX_CHAIN = 1 << 22;

function isCfb(bytes) {
    if (!bytes || bytes.length < 512) {
        return false;
    }
    for (let i = 0; i < SIGNATURE.length; i++) {
        if (bytes[i] !== SIGNATURE[i]) {
            return false;
        }
    }
    return true;
}

function readCfb(bytes) {
    if (!isCfb(bytes)) {
        throw new Error('The file is not an OLE2 compound document.');
    }
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const u32 = offset => view.getUint32(offset, true);
    const sectorSize = 1 << view.getUint16(0x1e, true);
    const miniSectorSize = 1 << view.getUint16(0x20, true);
    const fatSectors = u32(0x2c);
    const firstDirSector = u32(0x30);
    const miniCutoff = u32(0x38);
    const firstMiniFatSector = u32(0x3c);
    const miniFatSectors = u32(0x40);
    let difatSector = u32(0x44);
    const difatSectors = u32(0x48);

    if (sectorSize !== 512 && sectorSize !== 4096) {
        throw new Error('The compound document has an unsupported sector size.');
    }

    const sectorOffset = sector => (sector + 1) * sectorSize;

    // DIFAT: the first 109 FAT sector numbers live in the header, the rest in a chain of DIFAT sectors.
    const fatList = [];
    for (let i = 0; i < 109 && fatList.length < fatSectors; i++) {
        fatList.push(u32(0x4c + i * 4));
    }
    for (let n = 0; n < difatSectors && difatSector !== ENDOFCHAIN && difatSector !== FREESECT; n++) {
        const base = sectorOffset(difatSector);
        const perSector = sectorSize / 4 - 1;
        for (let i = 0; i < perSector && fatList.length < fatSectors; i++) {
            fatList.push(u32(base + i * 4));
        }
        difatSector = u32(base + perSector * 4);
    }

    const perFat = sectorSize / 4;
    const fat = new Uint32Array(fatList.length * perFat);
    fatList.forEach((sector, index) => {
        const base = sectorOffset(sector);
        for (let i = 0; i < perFat && base + i * 4 + 4 <= bytes.length; i++) {
            fat[index * perFat + i] = u32(base + i * 4);
        }
    });

    function chain(start, table) {
        const out = [];
        let sector = start;
        while (sector !== ENDOFCHAIN && sector !== FREESECT && sector < table.length) {
            out.push(sector);
            if (out.length > MAX_CHAIN) {
                throw new Error('The compound document has a looping sector chain.');
            }
            sector = table[sector];
        }
        return out;
    }

    function readChain(start, size) {
        const sectors = chain(start, fat);
        const out = new Uint8Array(size === undefined ? sectors.length * sectorSize : size);
        let written = 0;
        for (const sector of sectors) {
            if (written >= out.length) {
                break;
            }
            const base = sectorOffset(sector);
            const take = Math.min(sectorSize, out.length - written, bytes.length - base);
            if (take <= 0) {
                break;
            }
            out.set(bytes.subarray(base, base + take), written);
            written += take;
        }
        return out;
    }

    const dir = readChain(firstDirSector);
    const dirView = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
    const entries = [];
    for (let offset = 0; offset + 128 <= dir.length; offset += 128) {
        const nameLength = dirView.getUint16(offset + 0x40, true);
        const type = dir[offset + 0x42];
        let name = '';
        for (let i = 0; i + 2 < nameLength && i < 64; i += 2) {
            name += String.fromCharCode(dirView.getUint16(offset + i, true));
        }
        entries.push({
            name,
            type,
            start: dirView.getUint32(offset + 0x74, true),
            size: dirView.getUint32(offset + 0x78, true)
        });
    }
    const root = entries[0];
    if (!root || root.type !== 5) {
        throw new Error('The compound document has no root entry.');
    }

    const miniStream = readChain(root.start, root.size);
    let miniFat = new Uint32Array(0);
    if (miniFatSectors > 0 && firstMiniFatSector !== ENDOFCHAIN) {
        const raw = readChain(firstMiniFatSector);
        miniFat = new Uint32Array(raw.length / 4);
        const rawView = new DataView(raw.buffer, raw.byteOffset, raw.byteLength);
        for (let i = 0; i < miniFat.length; i++) {
            miniFat[i] = rawView.getUint32(i * 4, true);
        }
    }

    function readMini(start, size) {
        const out = new Uint8Array(size);
        let written = 0;
        for (const sector of chain(start, miniFat)) {
            if (written >= size) {
                break;
            }
            const base = sector * miniSectorSize;
            const take = Math.min(miniSectorSize, size - written, miniStream.length - base);
            if (take <= 0) {
                break;
            }
            out.set(miniStream.subarray(base, base + take), written);
            written += take;
        }
        return out;
    }

    const streams = new Map();
    for (const entry of entries) {
        if (entry.type === 2 && !streams.has(entry.name)) {
            streams.set(entry.name, entry);
        }
    }

    return {
        names: () => Array.from(streams.keys()),
        has: name => streams.has(name),
        // Streams below the cutoff live in the mini stream of the root entry.
        read(name) {
            const entry = streams.get(name);
            if (!entry) {
                return undefined;
            }
            return entry.size < miniCutoff ? readMini(entry.start, entry.size) : readChain(entry.start, entry.size);
        }
    };
}

module.exports = { isCfb, readCfb };
