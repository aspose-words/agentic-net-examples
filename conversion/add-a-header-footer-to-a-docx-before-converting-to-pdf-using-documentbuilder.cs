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

        // Add a primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Sample Header Text");

        // Add a primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Sample Footer Text");

        // Return to the main body and add some content.
        builder.MoveToSection(0);
        builder.Writeln("Document body content.");

        // Save the document as DOCX (the input file for conversion).
        const string docxPath = "input.docx";
        doc.Save(docxPath, SaveFormat.Docx);

        // Load the saved DOCX file.
        Document loadedDoc = new Document(docxPath);

        // Convert the loaded document to PDF.
        const string pdfPath = "output.pdf";
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF conversion failed: output file not found.");
    }
}
