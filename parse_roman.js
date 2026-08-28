function romanToInt(s) {
  const rom = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
  let num = 0;
  for (let i = 0; i < s.length; i++) {
    if (rom[s[i].toLowerCase()] < (rom[s[i + 1]?.toLowerCase()] || 0)) {
      num -= rom[s[i].toLowerCase()];
    } else {
      num += rom[s[i].toLowerCase()];
    }
  }
  return num;
}
function compareClass(a, b) {
  const getParts = (str) => {
    const match = str.trim().match(/^([IVXLCDMivxlcdm]+)(.*)$/);
    if (match) {
      return { num: romanToInt(match[1]), suffix: match[2].trim() };
    }
    return { num: 0, suffix: str.trim() };
  };
  const partA = getParts(a);
  const partB = getParts(b);
  if (partA.num !== partB.num) {
    return partA.num - partB.num;
  }
  return partA.suffix.localeCompare(partB.suffix, 'id-ID', { numeric: true });
}

console.log(['XII IPS', 'X IPA', 'XI IPA 2', 'XI IPA 1'].sort(compareClass));
