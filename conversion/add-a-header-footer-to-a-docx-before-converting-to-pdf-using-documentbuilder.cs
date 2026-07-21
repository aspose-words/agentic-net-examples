using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable a different header/footer for the first page (optional).
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Add a header to the first page.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Write("Sample Header");

        // Add a footer to the first page.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterFirst);
        builder.Write("Sample Footer");

        // Return to the main body and add some content.
        builder.MoveToSection(0);
        builder.Writeln("This is the body of the document.");

        // Save the document as DOCX – this will be the source for conversion.
        const string docxPath = "sample.docx";
        doc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX file and convert it to PDF.
        Document loadedDoc = new Document(docxPath);
        const string pdfPath = "output.pdf";
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF conversion failed.");
    }
}
