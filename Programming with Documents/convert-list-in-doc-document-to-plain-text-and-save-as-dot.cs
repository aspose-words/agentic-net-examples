using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input DOC file containing the list.
        string inputPath = @"C:\Docs\Source.doc";

        // Output DOT file that will contain the plain‑text representation.
        string outputPath = @"C:\Docs\Result.dot";

        // Load the source document.
        Document srcDoc = new Document(inputPath);

        // Update list labels so they are correct before conversion.
        srcDoc.UpdateListLabels();

        // Configure text save options to simplify list labels (ASCII symbols).
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            SimplifyListLabels = true
        };

        // Export the document to plain text using the configured options.
        string plainText = srcDoc.ToString(txtOptions);

        // Create a new blank document and insert the plain‑text content.
        Document destDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(destDoc);
        builder.Writeln(plainText);

        // Save the new document as a Word template (DOT format).
        destDoc.Save(outputPath, SaveFormat.Dot);
    }
}
