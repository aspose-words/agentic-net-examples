using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define output folder and file name.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "ReportWithConfidentialWatermark.docx");

        // Create a new blank document.
        Document doc = new Document();

        // (Optional) Add some sample content to simulate a report.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Automated Report");
        builder.Writeln("Generated on: " + DateTime.Now);
        builder.Writeln();
        builder.Writeln("This document contains the results of the automated reporting process.");

        // Configure text watermark options.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.Red,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false // Make the watermark fully opaque.
        };

        // Apply the confidential text watermark to the document.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Save the document with the watermark.
        doc.Save(outputPath);

        // Simple validation: confirm that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Report generated successfully with watermark at:");
            Console.WriteLine(outputPath);
        }
        else
        {
            Console.WriteLine("Failed to generate the report.");
        }
    }
}
