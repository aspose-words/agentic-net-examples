using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToPlainTextConverter
{
    static void Main()
    {
        // Path to the source DOC file that contains the list.
        const string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path where the resulting DOCX file will be saved.
        const string outputPath = @"C:\Docs\PlainTextDocument.docx";

        // Load the original document as plain text. The PlainTextDocument class extracts
        // the textual representation of the document, including list items.
        PlainTextDocument plainTextDoc = new PlainTextDocument(inputPath);
        string extractedText = plainTextDoc.Text;

        // Create a new blank Word document.
        Document newDoc = new Document();

        // Insert the extracted plain‑text into the new document.
        DocumentBuilder builder = new DocumentBuilder(newDoc);
        builder.Writeln(extractedText);

        // Save the new document as DOCX.
        newDoc.Save(outputPath, SaveFormat.Docx);
    }
}
