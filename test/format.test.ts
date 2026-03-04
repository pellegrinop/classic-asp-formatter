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
            // Now expecting same-level indentation
            assert.strictEqual(output, '<%\nline1\n\nline2\n%>');
        });

        it('should collapse multiple empty lines', () => {
            const input = '<%\nline1\n\n\n\nline2\n%>';
            const output = formatASP(input);
            // Now expecting same-level indentation
            assert.strictEqual(output, '<%\nline1\n\nline2\n%>');
        });
    });

    describe('Inline ASP and Indentation', () => {
        it('should handle <%= inline %> without adding newlines', () => {
            const input = '<input value="<%= mrEchoSession("mrecho") %>">';
            const output = formatASP(input);
            assert.strictEqual(output, '<input value="<%= mrEchoSession("mrecho") %>">');
        });

        it('should align closing tag with opening tag', () => {
            const input = '<div>\n    <%\n    Response.Write "test"\n    %>\n</div>';
            const output = formatASP(input);
            const expected = '<div>\n    <%\n    Response.Write "test"\n    %>\n</div>';
            assert.strictEqual(output, expected);
        });

        it('should handle multiline <%= ... %> correctly', () => {
            const input = 'value="<%\n    = mrEchoSession("mrecho")\n%>"';
            const output = formatASP(input);
            // Opening tag column is 7 (length of 'value="'), so content should have 7 spaces
            const expected = 'value="<%\n       = mrEchoSession("mrecho")\n       %>"';
            assert.strictEqual(output, expected);
        });
    });
});
