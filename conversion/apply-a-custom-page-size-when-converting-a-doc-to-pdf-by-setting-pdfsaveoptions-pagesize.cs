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
        builder.Writeln("Sample DOC content for custom page size.");
        source.Save("input.doc", SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document("input.doc");

        // Set a custom page size (420 x 595 points) for the first section.
        // Page dimensions are specified in points (1 point = 1/72 inch).
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 420f;
        section.PageSetup.PageHeight = 595f;

        // Save the document as PDF using the custom page size.
        doc.Save("output_custom_page.pdf", SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists("output_custom_page.pdf"))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
