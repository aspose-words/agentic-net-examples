using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with a custom page size.");

        // Define a custom page size (e.g., 5 inches wide by 7 inches tall).
        double widthInches = 5.0;
        double heightInches = 7.0;
        double widthPoints = ConvertUtil.InchToPoint(widthInches);
        double heightPoints = ConvertUtil.InchToPoint(heightInches);

        // Apply the custom size to the current section.
        builder.PageSetup.PaperSize = PaperSize.Custom;
        builder.PageSetup.PageWidth = widthPoints;
        builder.PageSetup.PageHeight = heightPoints;

        // Configure PDF save options (no PageSize property needed).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Determine the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomPageSize.pdf");

        // Save the document as PDF with the custom page size.
        doc.Save(outputPath, pdfOptions);

        // Validate that the PDF was created.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("Failed to create the PDF file.");
        }

        Console.WriteLine($"PDF successfully saved to: {outputPath}");
    }
}
