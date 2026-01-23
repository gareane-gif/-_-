
function normalizeId2(val) {
    if (val == null) return '';
    if (typeof val === 'number') return String(Math.floor(val));
    let s = String(val).trim();
    s = s.replace(/[\u0660-\u0669]/g, c => String(c.charCodeAt(0) - 0x0660));
    s = s.replace(/[\u06F0-\u06F9]/g, c => String(c.charCodeAt(0) - 0x06F0));
    s = s.replace(/\s+/g, '');
    s = s.replace(/[^a-zA-Z0-9]/g, '');
    if (/^\d+$/.test(s)) s = s.replace(/^0+/, '');
    return s.toUpperCase();
}

function isIdMatch(target, candidate) {
    const normTarget = normalizeId2(target);
    const normCandidate = normalizeId2(candidate);
    if (!normTarget || !normCandidate) return false;
    if (normTarget === normCandidate) return true;

    const hasLetters = /[A-Z]/.test(normTarget);
    if (hasLetters) {
        if (normTarget !== normCandidate) return false;
    }

    const numTarget = normTarget.replace(/[^0-9]/g, '');
    const numCandidate = normCandidate.replace(/[^0-9]/g, '');
    if (numTarget && numCandidate && numTarget === numCandidate && numTarget.length >= 4) {
        return true;
    }
    return false;
}

console.log("Testing SE241002 vs AC241002 (should be false):", isIdMatch('SE241002', 'AC241002'));
console.log("Testing SE241002 vs SE241002 (should be true):", isIdMatch('SE241002', 'SE241002'));
console.log("Testing 241002 vs SE241002 (should be true - numeric fallback):", isIdMatch('241002', 'SE241002'));
console.log("Testing SE241002 vs 241002 (should be false - target has letters):", isIdMatch('SE241002', '241002'));
