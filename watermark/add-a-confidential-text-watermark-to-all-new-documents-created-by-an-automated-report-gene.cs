using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create an output folder for the generated reports.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // Simulate generating several reports.
        for (int i = 1; i <= 3; i++)
        {
            // Create a new blank document.
            Document doc = new Document();

            // Add some sample content.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Report #{i}");
            builder.Writeln("This is an automatically generated report.");

            // Define watermark appearance.
            TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 48,
                Color = Color.Red,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = false
            };

            // Apply the confidential text watermark.
            doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

            // Save the document.
            string fileName = Path.Combine(outputFolder, $"Report_{i}.docx");
            doc.Save(fileName);
        }
    }
}
