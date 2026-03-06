using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the input DOC file and the output XPS file.
        string inputPath = "Input.doc";
        string outputPath = "Output.xps";

        // Load the original DOC document.
        Document sourceDoc = new Document(inputPath);

        // Configure text save options to simplify list labels (plain text representation).
        TxtSaveOptions txtOptions = new TxtSaveOptions();
        txtOptions.SimplifyListLabels = true;

        // Export the document content (including lists) to a plain‑text string.
        string plainText = sourceDoc.ToString(txtOptions);

        // Create a new blank document and insert the extracted plain text.
        Document plainDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(plainDoc);
        builder.Writeln(plainText);

        // Save the new document as XPS using default XpsSaveOptions.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        plainDoc.Save(outputPath, xpsOptions);
    }
}
