using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders.
        string inputFolder = "InputDocs";
        string outputFolder = "OutputDocs";

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample Word documents to demonstrate batch processing.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            var builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is sample document #{i}.");
            string samplePath = Path.Combine(inputFolder, $"Sample{i}.docx");
            sampleDoc.Save(samplePath);
        }

        // Define watermark appearance.
        var watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.LightGray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };
        string watermarkText = "CONFIDENTIAL";

        // Process each .docx file in the input folder.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(filePath);
            doc.Watermark.SetText(watermarkText, watermarkOptions);

            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }

        // Optional: verify that output files were created.
        int processedCount = Directory.GetFiles(outputFolder, "*.docx").Length;
        Console.WriteLine($"Watermark applied to {processedCount} document(s).");
    }
}
