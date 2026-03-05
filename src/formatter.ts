/**
 * Simple formatter for Classic ASP that respects HTML inside strings.
 * @param text The text to format.
 * @returns The formatted text.
 */
export function formatASP(text: string): string {
    let result = '';
    const indentSize = 4;
    const getIndent = (level: number) => ' '.repeat(level * indentSize);

    // Track global indent level for nested blocks across chips? 
    // User wants "same level" relative to tag, so we reset for each chip.

    // Regex to split by ASP blocks
    const parts = text.split(/(<%[\s\S]*?%>)/g);

    for (const part of parts) {
        if (part.startsWith('<%') && part.endsWith('%>')) {
            // Check if it's an expression tag <%=
            const isExpression = part.startsWith('<%=');
            const openingTag = isExpression ? '<%=' : '<%';
            const content = part.substring(openingTag.length, part.length - 2);
            
            const lines = content.split('\n');
            const isInitiallyMultiLine = lines.length > 1;
            const trimmedLines = lines.map(l => l.trim()).filter(l => l !== '');
            let activeOpeningTag = openingTag;
            if (!isExpression && trimmedLines.length > 0 && trimmedLines[0].startsWith('=')) {
                activeOpeningTag = '<%=';
            }
            
            // Track if single line for potential normalization later
            const isInitiallySingleLineShort = !isInitiallyMultiLine && content.trim().length < 80;
            const lastNewlineIndex = result.lastIndexOf('\n');
            const lastLinePrefix = lastNewlineIndex === -1 ? result : result.substring(lastNewlineIndex + 1);
            const baseIndentMatch = lastLinePrefix.match(/^\s*/);
            const baseIndent = baseIndentMatch ? baseIndentMatch[0] : '';

            // ASP Block (Multi-line or long)
            const formattedLines: string[] = [];
            let indentLevel = 0; // Reset for each chip to align with opening tag
            let inCase = false;
            let lastLineEmpty = false;
            
            const processedItems: string[] = [];
            for (const line of lines) {
                const trimmed = line.trim();
                
                // Handle empty lines: preserve at most one
                if (!trimmed) {
                    if (!lastLineEmpty) {
                        formattedLines.push('');
                        processedItems.push('');
                        lastLineEmpty = true;
                    }
                    continue;
                }
                lastLineEmpty = false;

                let isElseLine = false;
                if (isElse(trimmed)) {
                    indentLevel--;
                    isElseLine = true;
                } else if (isEndBlock(trimmed)) {
                    if (inCase && trimmed.toLowerCase().startsWith('end select')) {
                        inCase = false;
                    }
                    indentLevel--;
                }

                let extraCaseIndent = 0;
                if (isCase(trimmed)) {
                    inCase = true;
                } else if (inCase) {
                    extraCaseIndent = 1;
                }

                const processedLine = formatOperators(trimmed);
                processedItems.push(processedLine);
                formattedLines.push(baseIndent + getIndent(indentLevel + extraCaseIndent) + processedLine);

                if (isStartBlock(trimmed) || isElseLine) {
                    indentLevel++;
                }
            }
            
            // Cleanup leading/trailing empty lines
            while (formattedLines.length > 0 && formattedLines[0].trim() === '') {
                formattedLines.shift();
            }
            while (formattedLines.length > 0 && formattedLines[formattedLines.length - 1].trim() === '') {
                formattedLines.pop();
            }

            if (formattedLines.length === 1 && !isInitiallyMultiLine) {
                const line = processedItems.filter(li => li.trim() !== '')[0].trim(); 
                if (line.length < 80) {
                    let finalOpeningTag = activeOpeningTag;
                    let contentLine = line;
                    let separator = ' ';

                    if (activeOpeningTag === '<%=' && contentLine.startsWith('= ')) {
                        separator = '';
                    } else if (activeOpeningTag === '<%=' && contentLine.startsWith('=')) {
                        separator = '';
                    }

                    result += finalOpeningTag + separator + contentLine + ' %>';
                    continue;
                }
            }

            result += activeOpeningTag + '\n' + formattedLines.join('\n') + '\n' + baseIndent + '%>';
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
    // Split by VBScript comment if exists
    const commentIndex = line.indexOf("'");
    let codePart = line;
    let commentPart = '';
    
    if (commentIndex !== -1) {
        codePart = line.substring(0, commentIndex);
        commentPart = line.substring(commentIndex);
    }

    // Regex to find strings: "..." (with escaped "")
    // And operators: =, <>, >, <, >=, <=, +, -, *, /, &, \
    const regex = /("[^"]*(?:""[^"]*)*")|(<[=>]?|>=?|<>|[+\-*/&\\=])/g;
    
    const formattedCode = codePart.replace(regex, (match, stringLiteral, operator) => {
        if (stringLiteral) {
            return stringLiteral; // Return strings as-is
        }
        if (operator) {
            return ` ${operator.trim()} `;
        }
        return match;
    }).replace(/\s+/g, ' ').trim();

    return (formattedCode + ' ' + commentPart).trim();
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
    
    const isSingleLineIf = /^\bIf\b.*?\bThen\b\s+[^'\s]+/i.test(line);
    if (isSingleLineIf) {
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

export function isElse(line: string): boolean {
    return /^\bElse\b/i.test(line) || /^\bElseIf\b/i.test(line);
}
