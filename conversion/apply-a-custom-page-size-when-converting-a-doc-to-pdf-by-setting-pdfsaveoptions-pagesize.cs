using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC file with a custom page size (500x700 points).
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        // Set a custom page size. 1 point = 1/72 inch.
        builder.PageSetup.PaperSize = PaperSize.Custom;
        builder.PageSetup.PageWidth = 500f;   // width in points
        builder.PageSetup.PageHeight = 700f;  // height in points

        builder.Writeln("This document will be converted to PDF with a custom page size.");
        source.Save("input.doc", SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document("input.doc");

        // Configure PDF save options (no need to set PageSize here).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as PDF using the custom page size defined in the document.
        doc.Save("output.pdf", pdfOptions);

        // Verify that the PDF was created.
        if (!File.Exists("output.pdf"))
            throw new InvalidOperationException("The PDF file was not created.");

        // Clean up the temporary DOC file.
        if (File.Exists("input.doc"))
            File.Delete("input.doc");
    }
}
