using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToPlainTextTemplate
{
    static void Main()
    {
        // Path to the source DOC file that contains the list.
        const string sourceDocPath = @"C:\Docs\SourceDocument.doc";

        // Path for a temporary plain‑text file that will hold the converted list.
        const string tempTxtPath = @"C:\Docs\TempPlainText.txt";

        // Path for the final DOTX template file.
        const string outputDotxPath = @"C:\Docs\ResultTemplate.dotx";

        // Load the original DOC document.
        Document sourceDoc = new Document(sourceDocPath);

        // Configure TxtSaveOptions to simplify list labels (plain‑text representation).
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            SimplifyListLabels = true   // Convert complex list symbols to simple ASCII characters.
        };

        // Save the document as plain text using the configured options.
        sourceDoc.Save(tempTxtPath, txtOptions);

        // Load the generated plain‑text file back into a new Document object.
        Document plainTextDoc = new Document(tempTxtPath);

        // Save the plain‑text document as a DOTX template.
        plainTextDoc.Save(outputDotxPath, SaveFormat.Dotx);
    }
}
