// Minimal typing of the vendored SheetJS SSF (number format engine); only the parts the webview uses.
interface SsfOptions {
    date1904?: boolean;
    [option: string]: unknown;
}
interface Ssf {
    format(fmt: string, value: unknown, options?: SsfOptions): string;
    is_date(fmt: string): boolean;
    _split(fmt: string): string[];
    _table: Record<number, string | undefined>;
}
declare const SSF: Ssf;
export default SSF;
