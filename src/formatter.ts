/**
 * Simple formatter for Classic ASP that respects HTML inside strings.
 * @param text The text to format.
 * @returns The formatted text.
 */
export function formatASP(text: string): string {
    let result = '';
    let indentLevel = 0;
    const indentSize = 4;
    const getIndent = (level: number) => ' '.repeat(level * indentSize);

    // Regex to split by ASP blocks
    const parts = text.split(/(<%[\s\S]*?%>)/g);

    for (const part of parts) {
        if (part.startsWith('<%') && part.endsWith('%>')) {
            // ASP Block
            const content = part.substring(2, part.length - 2);
            const lines = content.split('\n');
            const formattedLines: string[] = [];
            let inCase = false;
            let lastLineEmpty = false;
            
            for (const line of lines) {
                const trimmed = line.trim();
                
                // Handle empty lines: preserve at most one
                if (!trimmed) {
                    if (!lastLineEmpty) {
                        formattedLines.push('');
                        lastLineEmpty = true;
                    }
                    continue;
                }
                lastLineEmpty = false;

                // Naive indentation: check if current line is an end block
                if (isEndBlock(trimmed)) {
                    if (inCase && trimmed.toLowerCase().startsWith('end select')) {
                        inCase = false;
                    }
                    indentLevel = Math.max(0, indentLevel - 1);
                }

                // Special handling for Case indentation
                let extraCaseIndent = 0;
                if (isCase(trimmed)) {
                    inCase = true;
                } else if (inCase) {
                    extraCaseIndent = 1;
                }

                // Apply operator spacing while respecting strings
                const processedLine = formatOperators(trimmed);

                // Add indentation
                formattedLines.push(getIndent(indentLevel + 1 + extraCaseIndent) + processedLine);

                // Increment indent for next line if current line starts a block
                if (isStartBlock(trimmed)) {
                    indentLevel++;
                }
            }
            
            // Cleanup leading/trailing empty lines in the block
            while (formattedLines.length > 0 && formattedLines[0] === '') {
                formattedLines.shift();
            }
            while (formattedLines.length > 0 && formattedLines[formattedLines.length - 1] === '') {
                formattedLines.pop();
            }

            result += '<%\n' + formattedLines.join('\n') + '\n' + getIndent(indentLevel) + '%>';
        } else {
            // HTML Block
            result += part;
        }
    }

    return result;
}

/**
 * Ensures operators have spaces around them, skipping strings.
 * @param line The line to process.
 * @returns The line with formatted operators.
 */
export function formatOperators(line: string): string {
    // Regex to find strings: "..." (with escaped "")
    // And operators: =, <>, >, <, >=, <=, +, -, *, /, &, \
    const regex = /("[^"]*(?:""[^"]*)*")|(<[=>]?|>=?|<>|[+\-*/&\\=])/g;
    
    return line.replace(regex, (match, stringLiteral, operator) => {
        if (stringLiteral) {
            return stringLiteral; // Return strings as-is
        }
        if (operator) {
            return ` ${operator.trim()} `;
        }
        return match;
    }).replace(/\s+/g, ' ').trim(); // Clean up double spaces created
}

export function isStartBlock(line: string): boolean {
    // Basic VBScript start blocks
    const startKeywords = [
        /^\bIf\b(?=.*?\bThen\b)(?!.*?\bEnd If\b)/i,
        /^\bFor\b/i, 
        /^\bWhile\b/i, 
        /^\bDo\b/i, 
        /^\bSelect Case\b/i, 
        /^\bFunction\b/i, 
        /^\bSub\b/i, 
        /^\bClass\b/i,
        /^\bWith\b/i
    ];
    
    const isSingleLineIf = /^\bIf\b.*?\bThen\b\s+[^']+/i.test(line);
    if (isSingleLineIf && !line.includes('\n')) {
        return false;
    }

    return startKeywords.some(regex => regex.test(line));
}

export function isEndBlock(line: string): boolean {
    const endKeywords = [
        /^\bEnd If\b/i, 
        /^\bNext\b/i, 
        /^\bWend\b/i, 
        /^\bLoop\b/i, 
        /^\bEnd Select\b/i, 
        /^\bEnd Function\b/i, 
        /^\bEnd Sub\b/i, 
        /^\bEnd Class\b/i,
        /^\bEnd With\b/i
    ];
    return endKeywords.some(regex => regex.test(line));
}

export function isCase(line: string): boolean {
    return /^\bCase\b/i.test(line);
}
