using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToPdfConverter
{
    static void Main()
    {
        // Path to the source DOC/DOCX file that contains the list.
        string sourcePath = "InputDocument.docx";

        // Path where the resulting PDF will be saved.
        string pdfPath = "ListAsPlainText.pdf";

        // Load the document as plain text, automatically handling any format.
        PlainTextDocument plainTextDoc = new PlainTextDocument(sourcePath);
        string extractedText = plainTextDoc.Text;

        // Create a new blank Word document.
        Document pdfDocument = new Document();

        // Insert the extracted plain‑text into the new document.
        DocumentBuilder builder = new DocumentBuilder(pdfDocument);
        builder.Writeln(extractedText);

        // Save the document as PDF.
        pdfDocument.Save(pdfPath, SaveFormat.Pdf);
    }
}
