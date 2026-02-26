using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToPlainTextTemplate
{
    static void Main()
    {
        // Path to the source DOC file that contains the list.
        const string sourceDocPath = @"C:\Docs\SourceList.doc";

        // Load the DOC document.
        Document sourceDoc = new Document(sourceDocPath);

        // Ensure list labels are up‑to‑date before extracting text.
        sourceDoc.UpdateListLabels();

        // Save the document as plain text, simplifying list labels for readability.
        // This uses the TxtSaveOptions rule that provides the SimplifyListLabels property.
        const string tempTxtPath = @"C:\Docs\TempList.txt";
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            SimplifyListLabels = true   // Convert complex list symbols to simple ASCII.
        };
        sourceDoc.Save(tempTxtPath, txtOptions);   // Save rule.

        // Load the plain‑text representation from the temporary file.
        // PlainTextDocument constructor rule is used here.
        PlainTextDocument plain = new PlainTextDocument(tempTxtPath);
        string plainText = plain.Text;   // PlainTextDocument.Text property rule.

        // Create a new blank document that will become the DOT template.
        Document templateDoc = new Document();   // Document() constructor rule.
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert the extracted plain‑text list into the new document.
        builder.Writeln(plainText);

        // Save the new document as a Word template (DOT) using the Save rule with SaveFormat.Dot.
        const string outputDotPath = @"C:\Docs\ListPlainTextTemplate.dot";
        templateDoc.Save(outputDotPath, SaveFormat.Dot);
    }
}
