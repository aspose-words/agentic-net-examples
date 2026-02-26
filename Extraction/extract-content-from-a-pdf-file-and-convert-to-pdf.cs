using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfExtractAndConvert
{
    static void Main()
    {
        // Path to the source PDF file.
        string inputPdfPath = "input.pdf";

        // Path to the resulting PDF file.
        string outputPdfPath = "output.pdf";

        // Load the source PDF into an Aspose.Words Document.
        Document sourceDoc = new Document(inputPdfPath);

        // Extract plain text from the PDF.
        PlainTextDocument plainTextDoc = new PlainTextDocument(inputPdfPath);
        string extractedText = plainTextDoc.Text;

        // Create a new blank document and insert the extracted text.
        Document newDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(newDoc);
        builder.Writeln(extractedText);

        // Save the new document as PDF using PdfSaveOptions.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        newDoc.Save(outputPdfPath, pdfOptions);
    }
}
