using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample content for the DOCX document.");
        sourceDoc.Save("input.docx", SaveFormat.Docx);

        // Step 2: Load the DOCX and add header and footer.
        Document doc = new Document("input.docx");
        DocumentBuilder headerFooterBuilder = new DocumentBuilder(doc);

        // Add header.
        headerFooterBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        headerFooterBuilder.Writeln("Document Header");

        // Add footer.
        headerFooterBuilder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        headerFooterBuilder.Writeln("Page Footer");

        // Step 3: Convert the document with header/footer to PDF.
        doc.Save("output.pdf", SaveFormat.Pdf);

        // Step 4: Validate that the PDF was created.
        if (!File.Exists("output.pdf"))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        // Optional: Clean up intermediate files (not required, but keeps the folder tidy).
        // File.Delete("input.docx");
    }
}
