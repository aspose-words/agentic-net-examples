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

        // Add a primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Sample Header");

        // Add a primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Sample Footer");

        // Return to the main body and add some content.
        builder.MoveToSection(0);
        builder.Writeln("Hello World!");

        // Save the document as DOCX (input for conversion).
        string docxPath = "sample.docx";
        doc.Save(docxPath, SaveFormat.Docx);

        // Load the saved DOCX.
        Document loadedDoc = new Document(docxPath);

        // Convert to PDF.
        string pdfPath = "output.pdf";
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("PDF conversion failed: output file not created or empty.");

        // Optional cleanup (comment out if you want to keep the files).
        // File.Delete(docxPath);
        // File.Delete(pdfPath);
    }
}
