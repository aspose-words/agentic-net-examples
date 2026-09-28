using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF file.
        Document sourcePdf = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourcePdf);
        builder.Writeln("Sample PDF content generated for conversion.");
        string pdfPath = "input.pdf";
        sourcePdf.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF with PdfLoadOptions (ignoring password protection).
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        // No password is set; the option will ignore password protection if none is required.
        Document loadedPdf = new Document(pdfPath, loadOptions);

        // Save the loaded document as DOCX.
        string docxPath = "output.docx";
        loadedPdf.Save(docxPath, SaveFormat.Docx);

        // Validate that the DOCX file was created.
        if (!File.Exists(docxPath))
        {
            throw new InvalidOperationException("Expected output DOCX was not created.");
        }

        // Optional: Clean up sample files (comment out if you want to inspect the files).
        // File.Delete(pdfPath);
        // File.Delete(docxPath);
    }
}
