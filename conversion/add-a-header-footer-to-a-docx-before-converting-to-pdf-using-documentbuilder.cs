using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class AddHeaderFooterAndConvert
{
    public static void Main()
    {
        // Paths for the temporary DOCX and final PDF files.
        const string docxPath = "sample.docx";
        const string pdfPath = "sample.pdf";

        // -------------------------------------------------
        // 1. Create a new blank document and add header/footer.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different headers/footers for the first page if desired.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Add a primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Sample Header Text");

        // Add a primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Sample Footer Text");

        // Return to the main body of the document.
        builder.MoveToSection(0);
        builder.Writeln("This is the body of the document.");

        // -------------------------------------------------
        // 2. Save the document as DOCX (input for conversion).
        // -------------------------------------------------
        doc.Save(docxPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 3. Load the saved DOCX and convert it to PDF.
        // -------------------------------------------------
        Document loadedDoc = new Document(docxPath);
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // -------------------------------------------------
        // 4. Validate that the PDF was created.
        // -------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF conversion failed: output file was not created.");

        // Optional: clean up temporary DOCX file.
        if (File.Exists(docxPath))
            File.Delete(docxPath);
    }
}
