using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Newtonsoft.Json;

public class WatermarkConfig
{
    public string Text { get; set; }
    public string FontFamily { get; set; }
    public float FontSize { get; set; }
    public string ColorName { get; set; }   // e.g., "Red" or "Blue"
    public bool IsSemitrasparent { get; set; } // false = opaque, true = semi‑transparent
}

public class Program
{
    public static void Main()
    {
        // Define file paths.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        string configPath = Path.Combine(artifactsDir, "watermarkConfig.json");
        string docPath1 = Path.Combine(artifactsDir, "Sample1.docx");
        string docPath2 = Path.Combine(artifactsDir, "Sample2.docx");
        string outPath1 = Path.Combine(artifactsDir, "Sample1.Watermarked.docx");
        string outPath2 = Path.Combine(artifactsDir, "Sample2.Watermarked.docx");

        // 1. Create a simple JSON configuration file if it does not exist.
        if (!File.Exists(configPath))
        {
            var defaultConfig = new WatermarkConfig
            {
                Text = "Confidential",
                FontFamily = "Arial",
                FontSize = 48,
                ColorName = "Red",
                IsSemitrasparent = false
            };
            File.WriteAllText(configPath, JsonConvert.SerializeObject(defaultConfig, Formatting.Indented));
        }

        // 2. Create two sample source documents.
        CreateSampleDocument(docPath1, "This is the first sample document.");
        CreateSampleDocument(docPath2, "This is the second sample document.");

        // 3. Load watermark configuration.
        WatermarkConfig config = JsonConvert.DeserializeObject<WatermarkConfig>(File.ReadAllText(configPath));

        // 4. Apply the watermark to each document.
        ApplyWatermark(docPath1, outPath1, config);
        ApplyWatermark(docPath2, outPath2, config);

        // 5. Simple validation – ensure output files exist.
        Console.WriteLine(File.Exists(outPath1)
            ? $"Watermarked file created: {outPath1}"
            : $"Failed to create: {outPath1}");
        Console.WriteLine(File.Exists(outPath2)
            ? $"Watermarked file created: {outPath2}"
            : $"Failed to create: {outPath2}");
    }

    private static void CreateSampleDocument(string filePath, string content)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln(content);
        doc.Save(filePath);
    }

    private static void ApplyWatermark(string inputPath, string outputPath, WatermarkConfig config)
    {
        var doc = new Document(inputPath);

        var options = new TextWatermarkOptions
        {
            FontFamily = config.FontFamily,
            FontSize = config.FontSize,
            Color = Color.FromName(config.ColorName),
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = config.IsSemitrasparent
        };

        doc.Watermark.SetText(config.Text, options);
        doc.Save(outputPath);
    }
}
