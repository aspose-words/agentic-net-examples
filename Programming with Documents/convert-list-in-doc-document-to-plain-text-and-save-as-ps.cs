using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertListToPlainTextAndSaveAsPs
{
    static void Main()
    {
        // Input DOC file containing the list.
        const string inputPath = "InputDocument.doc";

        // Output PS file.
        const string outputPath = "OutputDocument.ps";

        // Load the existing DOC document.
        Document sourceDoc = new Document(inputPath);

        // Convert the whole document to plain‑text representation.
        // This includes list items as they appear in the text.
        string plainText = sourceDoc.ToString(SaveFormat.Text);

        // Create a new blank document and write the plain text into it.
        Document plainDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(plainDoc);
        builder.Writeln(plainText);

        // Prepare PS save options (required to set the format explicitly).
        PsSaveOptions psOptions = new PsSaveOptions
        {
            SaveFormat = SaveFormat.Ps
        };

        // Save the plain‑text document as PostScript.
        plainDoc.Save(outputPath, psOptions);
    }
}
