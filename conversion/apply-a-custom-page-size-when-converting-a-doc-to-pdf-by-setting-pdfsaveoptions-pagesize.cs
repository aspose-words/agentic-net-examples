using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing; // Required for Aspose.Drawing types if needed

public class Program
{
    public static void Main()
    {
        // Create a sample DOC file.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOC content.");
        source.Save("input.doc", SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document("input.doc");

        // Define a custom page size (5 inches x 7 inches) in points.
        float pageWidth = (float)ConvertUtil.InchToPoint(5);
        float pageHeight = (float)ConvertUtil.InchToPoint(7);

        // Apply the custom size to each section of the document.
        foreach (Section section in doc.Sections)
        {
            section.PageSetup.PaperSize = PaperSize.Custom;
            section.PageSetup.PageWidth = pageWidth;
            section.PageSetup.PageHeight = pageHeight;
        }

        // Convert the DOC to PDF using the custom page size.
        doc.Save("output.pdf", SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists("output.pdf"))
            throw new InvalidOperationException("The PDF file was not created.");
    }
}
