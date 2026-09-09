using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC file.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("This is a sample document for PDF conversion with a custom page size.");

        // Save the source document as DOC.
        source.Save("input.doc", SaveFormat.Doc);

        // Load the created DOC file.
        Document doc = new Document("input.doc");

        // Set a custom page size (A4: 595 x 842 points) for the first section.
        // Page dimensions are measured in points (1 point = 1/72 inch).
        doc.FirstSection.PageSetup.PageWidth = 595f;
        doc.FirstSection.PageSetup.PageHeight = 842f;

        // Configure PDF save options (no need to set PageSize here).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Convert the DOC to PDF using the specified options.
        doc.Save("output.pdf", pdfOptions);

        // Verify that the PDF was created.
        if (!File.Exists("output.pdf"))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optional: Clean up temporary files.
        File.Delete("input.doc");
    }
}
