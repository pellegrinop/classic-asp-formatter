import * as vscode from 'vscode';
import { formatASP } from './formatter';

/**
 * Activates the extension.
 * @param context The extension context.
 */
export function activate(context: vscode.ExtensionContext) {
    console.log('Classic ASP extension is now active');
    vscode.window.showInformationMessage('Classic ASP extension activated!');

    // 1. Diagnostics (Syntax validation)
    const diagnosticCollection = vscode.languages.createDiagnosticCollection('classic-asp');
    context.subscriptions.push(diagnosticCollection);

    vscode.workspace.onDidChangeTextDocument(event => {
        validateDocument(event.document, diagnosticCollection);
    });

    vscode.workspace.onDidOpenTextDocument(document => {
        validateDocument(document, diagnosticCollection);
    });

    // 2. Formatter
    const formatter = vscode.languages.registerDocumentFormattingEditProvider('classic-asp', {
        provideDocumentFormattingEdits(document: vscode.TextDocument): vscode.TextEdit[] {
            const text = document.getText();
            const formattedText = formatASP(text);
            
            if (formattedText !== text) {
                const fullRange = new vscode.Range(
                    document.positionAt(0),
                    document.positionAt(text.length)
                );
                return [vscode.TextEdit.replace(fullRange, formattedText)];
            }
            
            return [];
        }
    });

    // 3. Debug Command
    const disposable = vscode.commands.registerCommand('classic-asp.check', () => {
        vscode.window.showInformationMessage('Classic ASP Formatter is active!');
    });
    context.subscriptions.push(disposable);

    context.subscriptions.push(formatter);
}

/**
 * Performs basic syntax validation.
 * @param document The document to validate.
 * @param collection The diagnostic collection to add errors to.
 */
function validateDocument(document: vscode.TextDocument, collection: vscode.DiagnosticCollection) {
    if (document.languageId !== 'classic-asp') {
        return;
    }

    const text = document.getText();
    const diagnostics: vscode.Diagnostic[] = [];

    // Check for unclosed <% blocks
    const openTags = (text.match(/<%/g) || []).length;
    const closeTags = (text.match(/%>/g) || []).length;

    if (openTags > closeTags) {
        const lastOpen = text.lastIndexOf('<%');
        const pos = document.positionAt(lastOpen);
        diagnostics.push(new vscode.Diagnostic(
            new vscode.Range(pos, pos.translate(0, 2)),
            'Unclosed ASP block (<%)',
            vscode.DiagnosticSeverity.Error
        ));
    }

    // 3. Check for basic VBScript block balance
    const blockPairs = [
        { start: /\bIf\b(?=.*?\bThen\b)(?!.*?\bEnd If\b)/ig, end: /\bEnd If\b/ig, name: 'If' },
        { start: /\bFunction\b/ig, end: /\bEnd Function\b/ig, name: 'Function' },
        { start: /\bSub\b/ig, end: /\bEnd Sub\b/ig, name: 'Sub' },
        { start: /\bFor\b/ig, end: /\bNext\b/ig, name: 'For' },
        { start: /\bWhile\b/ig, end: /\bWend\b/ig, name: 'While' },
        { start: /\bDo\b/ig, end: /\bLoop\b/ig, name: 'Do' }
    ];

    for (const pair of blockPairs) {
        let starts = 0;
        let match: RegExpExecArray | null;
        pair.start.lastIndex = 0;
        while ((match = pair.start.exec(text)) !== null) {
            const lineEnd = text.indexOf('\n', match.index);
            const line = text.substring(match.index, lineEnd !== -1 ? lineEnd : text.length);
            const isSingleLineIf = pair.name === 'If' && /^\bIf\b.*?\bThen\b\s+[^']+/i.test(line.trim());
            if (!isSingleLineIf) {
                starts++;
            }
        }

        const ends = (text.match(pair.end) || []).length;
        
        if (starts > ends) {
            diagnostics.push(new vscode.Diagnostic(
                new vscode.Range(0, 0, 0, 0),
                `Possible missing 'End ${pair.name}' (Starts: ${starts}, Ends: ${ends})`,
                vscode.DiagnosticSeverity.Warning
            ));
        }
    }

    collection.set(document.uri, diagnostics);
}

/**
 * Deactivates the extension.
 */
export function deactivate() {
    // Clean up resources if needed
}
