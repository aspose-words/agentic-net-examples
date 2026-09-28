using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Create a sample PDF file with two pages.
        // -----------------------------------------------------------------
        Document pdfSource = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfSource);
        builder.Writeln("First page of the PDF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page of the PDF.");

        string pdfPath = "sample.pdf";
        pdfSource.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Load the PDF. Aspose.Words automatically ignores non‑critical
        // errors during PDF import, so no explicit IgnoreErrors flag is needed.
        // -----------------------------------------------------------------
        PdfLoadOptions loadOptions = new PdfLoadOptions(); // No IgnoreErrors property in current API.
        Document loadedPdf = new Document(pdfPath, loadOptions);

        // -----------------------------------------------------------------
        // Save the loaded document as DOCX. Page breaks are preserved by default.
        // -----------------------------------------------------------------
        string docxPath = "output.docx";
        loadedPdf.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Validate that the DOCX file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(docxPath))
        {
            throw new InvalidOperationException("The DOCX file was not created as expected.");
        }

        Console.WriteLine("Conversion completed successfully.");
    }
}
