using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToRtfConverter
{
    static void Main()
    {
        // Path to the source DOC file that contains the list.
        string sourcePath = @"C:\Docs\SourceDocument.doc";

        // Path where the resulting RTF file will be saved.
        string outputPath = @"C:\Docs\PlainTextOutput.rtf";

        // Load the original DOC document.
        Document sourceDoc = new Document(sourcePath);

        // Ensure that list labels are up‑to‑date before extracting plain text.
        sourceDoc.UpdateListLabels();

        // Extract the plain‑text representation of the document (including list items).
        PlainTextDocument plainText = new PlainTextDocument(sourcePath);
        string textContent = plainText.Text;

        // Create a new blank document that will hold the plain text.
        Document rtfDoc = new Document();

        // Insert the extracted text into the new document.
        DocumentBuilder builder = new DocumentBuilder(rtfDoc);
        builder.Writeln(textContent);

        // Save the new document in RTF format.
        rtfDoc.Save(outputPath, SaveFormat.Rtf);
    }
}
