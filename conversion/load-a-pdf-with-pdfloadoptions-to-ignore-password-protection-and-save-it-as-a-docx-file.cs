using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF file.
        Document samplePdf = new Document();
        DocumentBuilder builder = new DocumentBuilder(samplePdf);
        builder.Writeln("This is a sample PDF created for conversion.");
        const string pdfPath = "sample.pdf";
        samplePdf.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF using PdfLoadOptions (no password provided, effectively ignoring password protection).
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document loadedDoc = new Document(pdfPath, loadOptions);

        // Save the loaded document as DOCX.
        const string docxPath = "output.docx";
        loadedDoc.Save(docxPath, SaveFormat.Docx);

        // Validate that the DOCX file was created.
        if (!File.Exists(docxPath))
        {
            throw new InvalidOperationException("The DOCX file was not created as expected.");
        }

        // Optional: clean up generated files (comment out if you want to keep them).
        // File.Delete(pdfPath);
        // File.Delete(docxPath);
    }
}
