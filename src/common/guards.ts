export function isRecord(value: unknown): value is Record<string, unknown> {
    return typeof value === 'object' && value !== null;
}

export function errorMessage(reason: unknown): string {
    if (isRecord(reason) && typeof reason.message === 'string' && reason.message) {
        return reason.message;
    }
    return String(reason);
}
