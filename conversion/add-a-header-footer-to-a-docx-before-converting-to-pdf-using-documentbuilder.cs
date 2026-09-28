using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        // Add some body content.
        builder.Writeln("This is the main document body.");

        // Move to the primary header (creates it if it does not exist) and write text.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Sample Header Text");

        // Move to the primary footer (creates it if it does not exist) and write page numbers.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Page ");
        builder.InsertField("PAGE", "?");
        builder.Write(" of ");
        builder.InsertField("NUMPAGES", "?");

        // Save the document as DOCX.
        string docxPath = "sample.docx";
        source.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX file.
        Document doc = new Document(docxPath);

        // Convert the loaded document to PDF.
        string pdfPath = "output.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
