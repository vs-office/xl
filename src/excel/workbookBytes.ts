import { isRecord } from '../common/guards';

const MAX_FILE_BYTES = 1024 * 1024 * 1024;

function isArrayLike(value: unknown): value is ArrayLike<number> {
    return (
        isRecord(value) &&
        typeof value.length === 'number' &&
        Number.isInteger(value.length) &&
        value.length > 0 &&
        value.length <= MAX_FILE_BYTES
    );
}

export function toWorkbookBytes(body: unknown): Uint8Array | undefined {
    let out: Uint8Array | undefined;
    if (body instanceof Uint8Array) {
        out = body;
    } else if (body instanceof ArrayBuffer) {
        out = new Uint8Array(body);
    } else if (ArrayBuffer.isView(body)) {
        out = new Uint8Array(body.buffer, body.byteOffset, body.byteLength);
    } else if (isArrayLike(body)) {
        out = Uint8Array.from(body); // structured-clone fallback for array-likes
    }
    if (!out || out.length < 4 || out.length > MAX_FILE_BYTES || out[0] !== 0x50 || out[1] !== 0x4b) {
        return undefined;
    }
    return out;
}
