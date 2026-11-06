using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the main body of the document.");
        sourceDoc.Save("input.docx", SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document("input.docx");

        // Add a header.
        DocumentBuilder headerBuilder = new DocumentBuilder(doc);
        headerBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        headerBuilder.Writeln("Sample Header");

        // Add a footer.
        DocumentBuilder footerBuilder = new DocumentBuilder(doc);
        footerBuilder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        footerBuilder.Writeln("Sample Footer");

        // Convert the document to PDF.
        doc.Save("output.pdf", SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists("output.pdf"))
            throw new InvalidOperationException("Expected output PDF was not created.");
    }
}
