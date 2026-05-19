using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Drawing;
using Aspose.Words;

public class Program
{
    // Model that matches the JSON configuration file.
    private class WatermarkConfig
    {
        public string Text { get; set; } = "Sample Watermark";
        public string FontFamily { get; set; } = "Arial";
        public float FontSize { get; set; } = 36;
        public string Color { get; set; } = "#808080"; // Gray
        public bool IsSemitrasparent { get; set; } = true;
    }

    // Loads the configuration from a JSON file; creates a default file if it does not exist.
    private static WatermarkConfig LoadConfig(string path)
    {
        if (!File.Exists(path))
        {
            var defaultConfig = new WatermarkConfig();
            var json = JsonSerializer.Serialize(defaultConfig, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
            return defaultConfig;
        }

        var jsonText = File.ReadAllText(path);
        return JsonSerializer.Deserialize<WatermarkConfig>(jsonText) ?? new WatermarkConfig();
    }

    // Converts a hex colour string (e.g. "#FF0000") to a System.Drawing.Color.
    private static Color ParseColor(string hex)
    {
        try
        {
            return ColorTranslator.FromHtml(hex);
        }
        catch
        {
            return Color.Black;
        }
    }

    // Creates a simple document containing a single paragraph with the supplied text.
    private static Document CreateSampleDocument(string content)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln(content);
        return doc;
    }

    public static void Main()
    {
        const string configPath = "watermarkConfig.json";

        // Load configuration (or create a default one).
        WatermarkConfig config = LoadConfig(configPath);

        // Prepare watermark options based on the configuration.
        var watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = config.FontFamily,
            FontSize = config.FontSize,
            Color = ParseColor(config.Color),
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = config.IsSemitrasparent
        };

        // Documents to be generated.
        var documentsInfo = new List<(string fileName, string content)>
        {
            ("Document1.docx", "First sample document."),
            ("Document2.docx", "Second sample document."),
            ("Document3.docx", "Third sample document.")
        };

        foreach (var (fileName, content) in documentsInfo)
        {
            // Create a fresh document.
            Document doc = CreateSampleDocument(content);

            // Apply the text watermark using the configuration.
            doc.Watermark.SetText(config.Text, watermarkOptions);

            // Save the watermarked document.
            doc.Save(fileName);
        }

        // Simple validation: report which files were created.
        foreach (var (fileName, _) in documentsInfo)
        {
            Console.WriteLine(File.Exists(fileName)
                ? $"Created: {fileName}"
                : $"Failed to create: {fileName}");
        }
    }
}
