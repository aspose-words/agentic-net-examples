using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class PdfToDocxConversion
{
    public static void Main()
    {
        // Create a sample PDF file.
        Document samplePdf = new Document();
        DocumentBuilder builder = new DocumentBuilder(samplePdf);
        builder.Writeln("Sample PDF content with layout preservation.");
        string pdfPath = "sample.pdf";
        samplePdf.Save(pdfPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The sample PDF file was not created.");

        // Load the PDF with layout preservation options.
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        // Layout preservation is enabled by default; additional options can be set here if needed.

        Document pdfDocument = new Document(pdfPath, loadOptions);

        // Convert and save as DOCX.
        string docxPath = "converted.docx";
        pdfDocument.Save(docxPath, SaveFormat.Docx);

        // Validate that the DOCX file was created.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("The DOCX file was not created after conversion.");

        // Clean up temporary files (optional).
        File.Delete(pdfPath);
        // The converted DOCX remains for verification.
    }
}
