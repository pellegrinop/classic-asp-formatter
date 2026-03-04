import * as assert from 'assert';
import { formatASP, formatOperators } from '../src/formatter';

describe('Formatter Unit Tests', () => {
    
    describe('Operator Spacing', () => {
        it('should add spaces around =', () => {
            assert.strictEqual(formatOperators('x=y'), 'x = y');
        });

        it('should add spaces around &', () => {
            assert.strictEqual(formatOperators('a&b'), 'a & b');
        });

        it('should not add spaces inside strings', () => {
            const input = 'msg = "1+1=2"';
            assert.strictEqual(formatOperators(input), 'msg = "1+1=2"');
        });

        it('should handle complex expressions', () => {
            assert.strictEqual(formatOperators('testOp=1+2*3/4'), 'testOp = 1 + 2 * 3 / 4');
        });
    });

    describe('Select Case Indentation', () => {
        it('should indent content after Case', () => {
            const input = '<%\nSelect Case x\nCase 1\nResponse.Write "1"\nEnd Select\n%>';
            const output = formatASP(input);
            assert.ok(output.includes('    Case 1'));
            assert.ok(output.includes('        Response.Write "1"'));
        });
    });

    describe('Empty Line Preservation', () => {
        it('should preserve single empty lines', () => {
            const input = '<%\nline1\n\nline2\n%>';
            const output = formatASP(input);
            assert.ok(output.includes('line1\n\n    line2'));
        });

        it('should collapse multiple empty lines', () => {
            const input = '<%\nline1\n\n\n\nline2\n%>';
            const output = formatASP(input);
            // Should have exactly one empty line between line1 and line2
            const parts = output.split('\n');
            const emptyLines = parts.filter(p => p.trim() === '').length;
            // Expected: 1 between line1 and line2
            assert.strictEqual(emptyLines, 1); 
        });
    });
});
