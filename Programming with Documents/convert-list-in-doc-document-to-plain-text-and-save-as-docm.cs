using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToPlainTextConverter
{
    static void Main()
    {
        // Path to the source DOC file that contains the list.
        string inputPath = "input.doc";

        // Path where the resulting DOCM file will be saved.
        string outputPath = "output.docm";

        // Load the existing DOC document.
        Document sourceDoc = new Document(inputPath);

        // Configure text save options to simplify list labels when converting to plain text.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            SimplifyListLabels = true
        };

        // Export the document content to plain text using the configured options.
        string plainText = sourceDoc.ToString(txtOptions);

        // Create a new blank document to hold the plain‑text representation.
        Document resultDoc = new Document();

        // Insert the plain text into the new document.
        DocumentBuilder builder = new DocumentBuilder(resultDoc);
        builder.Writeln(plainText);

        // Save the result as a macro‑enabled Word document (DOCM).
        resultDoc.Save(outputPath, SaveFormat.Docm);
    }
}
