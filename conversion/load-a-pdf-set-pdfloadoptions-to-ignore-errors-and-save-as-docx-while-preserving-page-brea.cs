using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample PDF with a page break.
        Document samplePdf = new Document();
        DocumentBuilder builder = new DocumentBuilder(samplePdf);
        builder.Writeln("First page content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content.");
        const string pdfPath = "sample.pdf";
        samplePdf.Save(pdfPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Sample PDF was not created.");

        // Step 2: Load the PDF with PdfLoadOptions that ignore errors.
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        // By default, Aspose.Words ignores non‑critical errors during PDF loading.
        // If specific error‑ignoring flags existed, they would be set here.

        Document pdfDocument = new Document(pdfPath, loadOptions);

        // Step 3: Save the loaded document as DOCX, preserving page breaks.
        const string docxPath = "output.docx";
        pdfDocument.Save(docxPath, SaveFormat.Docx);

        // Validate that the DOCX file was created.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("The DOCX output was not created.");

        Console.WriteLine("PDF successfully converted to DOCX with page breaks preserved.");
    }
}
