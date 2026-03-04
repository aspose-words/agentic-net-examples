using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class PdfExtractAndConvert
{
    static void Main()
    {
        // Path to the source PDF file.
        string sourcePdfPath = "input.pdf";

        // Load the PDF document. PdfLoadOptions can be used if specific loading behavior is required.
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document sourceDoc = new Document(sourcePdfPath, loadOptions);

        // Extract plain text from the PDF using PlainTextDocument.
        PlainTextDocument plainText = new PlainTextDocument(sourcePdfPath);
        string extractedText = plainText.Text;

        // Create a new blank Word document.
        Document outputDoc = new Document();

        // Insert the extracted text into the new document.
        DocumentBuilder builder = new DocumentBuilder(outputDoc);
        builder.Writeln(extractedText);

        // Prepare PDF save options (default options are sufficient for a basic conversion).
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Save the new document as a PDF.
        outputDoc.Save("output.pdf", saveOptions);
    }
}
