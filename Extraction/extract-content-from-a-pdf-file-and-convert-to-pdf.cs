using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfExtractAndConvert
{
    static void Main()
    {
        // Path to the source PDF file.
        string sourcePdfPath = @"C:\Input\source.pdf";

        // Path to the resulting PDF file.
        string outputPdfPath = @"C:\Output\extracted.pdf";

        // Load the PDF document into an Aspose.Words Document.
        Document sourceDoc = new Document(sourcePdfPath);

        // Extract all textual content from the loaded PDF.
        string extractedText = sourceDoc.GetText();

        // Create a new blank Word document.
        Document newDoc = new Document();

        // Insert the extracted text into the new document.
        DocumentBuilder builder = new DocumentBuilder(newDoc);
        builder.Writeln(extractedText);

        // Save the new document as a PDF file.
        newDoc.Save(outputPdfPath, SaveFormat.Pdf);
    }
}
