using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToPlainTextWordml
{
    static void Main()
    {
        // Path to the source DOC document that contains a list.
        string sourcePath = @"C:\Docs\SourceDocument.doc";

        // Load the DOC document.
        Document sourceDoc = new Document(sourcePath);

        // Ensure that list labels are up‑to‑date before extracting text.
        sourceDoc.UpdateListLabels();

        // Extract the whole document as plain text (lists are rendered as text).
        string plainText = sourceDoc.ToString(SaveFormat.Text);

        // Create a new blank document.
        Document plainDoc = new Document();

        // Insert the extracted plain text into the new document.
        DocumentBuilder builder = new DocumentBuilder(plainDoc);
        builder.Writeln(plainText);

        // Save the new document in WordML (Microsoft Word 2003 XML) format.
        string outputPath = @"C:\Docs\PlainTextDocument.xml";
        plainDoc.Save(outputPath, SaveFormat.WordML);
    }
}
