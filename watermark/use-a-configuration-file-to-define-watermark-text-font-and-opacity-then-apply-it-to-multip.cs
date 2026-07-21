using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class WatermarkConfig
{
    public string Text { get; set; }
    public string FontFamily { get; set; }
    public float FontSize { get; set; }
    public string ColorHex { get; set; }   // e.g., "#FF0000"
    public bool IsSemitrasparent { get; set; }
}

public class Program
{
    private const string ConfigFileName = "watermarkConfig.json";
    private const string OutputFolder = "Output";

    public static void Main()
    {
        // Ensure output directory exists.
        Directory.CreateDirectory(OutputFolder);

        // Create a sample configuration file if it does not exist.
        if (!File.Exists(ConfigFileName))
        {
            var defaultConfig = new WatermarkConfig
            {
                Text = "Confidential",
                FontFamily = "Arial",
                FontSize = 48,
                ColorHex = "#FF0000",          // Red
                IsSemitrasparent = false      // Opaque
            };
            var json = JsonSerializer.Serialize(defaultConfig, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(ConfigFileName, json);
        }

        // Load configuration.
        var configJson = File.ReadAllText(ConfigFileName);
        var config = JsonSerializer.Deserialize<WatermarkConfig>(configJson);

        // Prepare a list of sample documents.
        var sourceDocs = CreateSampleDocuments();

        // Apply watermark to each document.
        int index = 1;
        foreach (var srcPath in sourceDocs)
        {
            var doc = new Document(srcPath);

            var options = new TextWatermarkOptions
            {
                FontFamily = config.FontFamily,
                FontSize = config.FontSize,
                Color = ColorTranslator.FromHtml(config.ColorHex),
                IsSemitrasparent = config.IsSemitrasparent,
                Layout = WatermarkLayout.Diagonal
            };

            doc.Watermark.SetText(config.Text, options);

            string outPath = Path.Combine(OutputFolder, $"Watermarked_{index}.docx");
            doc.Save(outPath);

            // Simple validation: ensure the file was created.
            if (File.Exists(outPath))
                Console.WriteLine($"Document {outPath} saved successfully.");
            else
                Console.WriteLine($"Failed to save document {outPath}.");

            index++;
        }
    }

    // Creates a few simple documents with placeholder text and returns their file paths.
    private static List<string> CreateSampleDocuments()
    {
        var paths = new List<string>();
        for (int i = 1; i <= 3; i++)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample document {i}");
            builder.Writeln("This is a test document used to demonstrate watermark application.");
            string path = $"Sample_{i}.docx";
            doc.Save(path);
            paths.Add(path);
        }
        return paths;
    }
}
