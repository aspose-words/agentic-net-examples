using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToPlainTextWordml
{
    static void Main()
    {
        // Input DOC file that contains the list.
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Output WORDML file (Word 2003 XML format).
        string outputPath = @"C:\Docs\OutputDocument.xml";

        // Load the source document.
        Document sourceDoc = new Document(inputPath);

        // Ensure list labels are up‑to‑date so they appear in the plain‑text export.
        sourceDoc.UpdateListLabels();

        // Export the whole document to plain text (including list labels).
        // Using SaveFormat.Text preserves the list formatting in the text.
        string plainText = sourceDoc.ToString(SaveFormat.Text);

        // Create a new blank document.
        Document plainTextDoc = new Document();

        // Write the extracted plain text into the new document.
        DocumentBuilder builder = new DocumentBuilder(plainTextDoc);
        builder.Writeln(plainText);

        // Save the new document as WORDML (Word 2003 XML).
        plainTextDoc.Save(outputPath, SaveFormat.WordML);
    }
}
