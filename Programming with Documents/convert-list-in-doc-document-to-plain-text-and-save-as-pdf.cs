using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ListToPlainTextPdf
{
    static void Main()
    {
        // Path to the source DOC/DOCX file that contains the list.
        string inputFile = @"C:\Docs\SourceListDocument.docx";

        // Path where the resulting PDF will be saved.
        string outputFile = @"C:\Docs\PlainTextList.pdf";

        // Load the source document as plain text. The PlainTextDocument class extracts
        // the textual content of the file, ignoring all formatting and list labels.
        PlainTextDocument plainTextDoc = new PlainTextDocument(inputFile);
        string extractedText = plainTextDoc.Text;

        // Create a new blank Word document.
        Document pdfDocument = new Document();

        // Use DocumentBuilder to insert the extracted plain text into the new document.
        DocumentBuilder builder = new DocumentBuilder(pdfDocument);
        builder.Writeln(extractedText);

        // Save the new document as PDF.
        pdfDocument.Save(outputFile, SaveFormat.Pdf);
    }
}
