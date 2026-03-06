using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToPlainTextConverter
{
    static void Main()
    {
        // Path to the source DOC file that contains the list.
        string sourceDocPath = "input.doc";

        // Path where the resulting DOCX file will be saved.
        string resultDocxPath = "output.docx";

        // Load the source document as plain text. The PlainTextDocument class extracts
        // the textual representation of the document, including list items.
        PlainTextDocument plainText = new PlainTextDocument(sourceDocPath);
        string extractedText = plainText.Text;

        // Create a new blank Word document.
        Document newDoc = new Document();

        // Use DocumentBuilder to insert the extracted plain‑text into the new document.
        DocumentBuilder builder = new DocumentBuilder(newDoc);
        builder.Writeln(extractedText);

        // Save the new document as DOCX.
        newDoc.Save(resultDocxPath, SaveFormat.Docx);
    }
}
