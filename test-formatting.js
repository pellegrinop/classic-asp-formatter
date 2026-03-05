const fs = require('fs');
const { formatASP } = require('./out/src/formatter.js');

const text = fs.readFileSync('test.asp', 'utf-8');
try {
    const formatted = formatASP(text);
    console.log("Formatted successfully! Output length:", formatted.length);
    console.log("Differs from original:", formatted !== text);
} catch (e) {
    console.error("Format error:", e);
}
