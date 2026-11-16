import { env, Uri, window } from 'vscode';

const ALLOWED = ['http:', 'https:', 'mailto:'];

export function isSafeLink(url: string): boolean {
    const scheme = /^([a-zA-Z][a-zA-Z0-9+.-]*:)/.exec(url.trim())?.[1];
    return scheme !== undefined && ALLOWED.includes(scheme.toLowerCase()) && url.length < 4096;
}

export function openExternalLink(url: string) {
    if (!isSafeLink(url)) {
        void window.showWarningMessage('xl only opens http, https and mailto links.');
        return;
    }
    void env.openExternal(Uri.parse(url.trim()));
}
