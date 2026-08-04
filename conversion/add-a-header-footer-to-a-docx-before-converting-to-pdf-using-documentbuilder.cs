using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string docxPath = "sample.docx";
        string pdfPath = "sample.pdf";

        // -----------------------------------------------------------------
        // 1. Create a new blank document and add a header and a footer.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure the same header/footer appears on every page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = false;
        builder.PageSetup.OddAndEvenPagesHeaderFooter = false;

        // Add header text.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Sample Header Text");

        // Add footer text.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Sample Footer Text");

        // Add some body content so the PDF is not empty.
        builder.MoveToSection(0);
        builder.Writeln("This is the body of the document.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content.");

        // Save the document as DOCX (input file).
        doc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the saved DOCX and convert it to PDF.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docxPath);
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 3. Verify that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optional cleanup (comment out if you want to inspect the files).
        // File.Delete(docxPath);
        // File.Delete(pdfPath);
    }
}
