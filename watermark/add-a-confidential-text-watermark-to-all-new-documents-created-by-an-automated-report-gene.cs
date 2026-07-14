using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define output folder and file.
        string outputFolder = "Output";
        Directory.CreateDirectory(outputFolder);
        string outputPath = Path.Combine(outputFolder, "ReportWithWatermark.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample content to the report.
        builder.Writeln("Automated Report");
        builder.Writeln("Generated on: " + DateTime.Now);

        // Configure text watermark options.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Red,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Apply the confidential text watermark.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Save the document.
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Report saved with watermark at: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the report.");
        }
    }
}
