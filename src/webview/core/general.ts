// ---------- General ----------
function incDigits(D: string, keep: number): [string, number] {
    // round digit string D to `keep` sig digits (half up); returns [digits, carry]
    if (keep >= D.length) return [D, 0];
    const up = D.charCodeAt(keep) >= 53;
    const head = D.slice(0, keep);
    if (!up) return [head, 0];
    const a = head.split('');
    let i = keep - 1;
    while (i >= 0) {
        if (a[i] === '9') {
            a[i] = '0';
            i--;
        } else {
            a[i] = String.fromCharCode(a[i].charCodeAt(0) + 1);
            break;
        }
    }
    if (i < 0) return ['1' + a.join(''), 1];
    return [a.join(''), 0];
}
function stripZeros(D: string) {
    let n = D.length;
    while (n > 1 && D.charCodeAt(n - 1) === 48) n--;
    return n === D.length ? D : D.slice(0, n);
}
// Excel General: <= 11 chars, 15 significant digits, scientific outside [1e-4, 1e11)
function generalSlow(v: number): string {
    const neg = v < 0,
        a = neg ? -v : v;
    const es = a.toExponential(14);
    const k = es.indexOf('e');
    let ex = +es.slice(k + 1);
    let D = es.charAt(0) + es.slice(2, k);
    let out: string;
    if (ex >= -4 && ex <= 10) {
        const keep = ex >= 0 ? (ex === 10 ? 11 : 10) : 10 + ex;
        const r = incDigits(D, keep);
        D = r[0];
        ex += r[1];
        if (ex <= 10) {
            D = stripZeros(D);
            if (ex >= 0) {
                out =
                    D.length <= ex + 1 ? D + '0'.repeat(ex + 1 - D.length) : D.slice(0, ex + 1) + '.' + D.slice(ex + 1);
            } else out = '0.' + '0'.repeat(-ex - 1) + D;
            return neg ? '-' + out : out;
        }
        D = es.charAt(0) + es.slice(2, k);
        ex = +es.slice(k + 1);
    }
    const ae = ex < 0 ? -ex : ex;
    const keep = 1 + 7 - (ae >= 100 ? 3 : 2);
    const r = incDigits(D, keep);
    D = stripZeros(r[0]);
    ex += r[1];
    const ae2 = ex < 0 ? -ex : ex;
    out =
        D.charAt(0) +
        (D.length > 1 ? '.' + D.slice(1) : '') +
        'E' +
        (ex < 0 ? '-' : '+') +
        (ae2 < 10 ? '0' + ae2 : ae2);
    return neg ? '-' + out : out;
}
function generalNum(v: number): string {
    if ((v | 0) === v) return '' + v;
    if (v !== v || v === Infinity || v === -Infinity) return '#NUM!';
    const a = v < 0 ? -v : v;
    if (a >= 1e-4 && a < 1e11) {
        const s = '' + v;
        if (s.length - (v < 0 ? 1 : 0) <= 11 && s.indexOf('e') < 0) return s;
    }
    return generalSlow(v);
}

export { generalNum };
