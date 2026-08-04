using System;
using System.IO;
using System.Text.Json;
using System.Drawing;
using Aspose.Words;

public class Program
{
    // Model that matches the JSON configuration file.
    private class WatermarkConfig
    {
        public string Text { get; set; }
        public string FontFamily { get; set; }
        public float FontSize { get; set; }
        public string Color { get; set; }               // Color name or HTML hex (e.g., "#FF0000")
        public bool IsSemitrasparent { get; set; }      // true = semi‑transparent, false = opaque
    }

    public static void Main()
    {
        // 1. Create a default configuration file if it does not exist.
        const string configFile = "watermarkConfig.json";
        if (!File.Exists(configFile))
        {
            var defaultConfig = new WatermarkConfig
            {
                Text = "CONFIDENTIAL",
                FontFamily = "Arial",
                FontSize = 48,
                Color = "#FF0000",          // Red
                IsSemitrasparent = false   // Fully opaque
            };
            string json = JsonSerializer.Serialize(defaultConfig, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(configFile, json);
        }

        // 2. Load configuration from the JSON file.
        WatermarkConfig config = JsonSerializer.Deserialize<WatermarkConfig>(File.ReadAllText(configFile));

        // 3. Create sample source documents (if they do not already exist).
        string[] sourceDocs = { "Sample1.docx", "Sample2.docx", "Sample3.docx" };
        for (int i = 0; i < sourceDocs.Length; i++)
        {
            if (!File.Exists(sourceDocs[i]))
            {
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln($"This is the content of {Path.GetFileNameWithoutExtension(sourceDocs[i])}.");
                doc.Save(sourceDocs[i]);
            }
        }

        // 4. Apply the watermark to each document using the loaded configuration.
        foreach (string srcPath in sourceDocs)
        {
            // Load the document.
            Document doc = new Document(srcPath);

            // Build watermark options from the configuration.
            TextWatermarkOptions options = new TextWatermarkOptions
            {
                FontFamily = config.FontFamily,
                FontSize = config.FontSize,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = config.IsSemitrasparent
            };

            // Parse the color (supports named colors or HTML hex).
            try
            {
                options.Color = ColorTranslator.FromHtml(config.Color);
            }
            catch
            {
                // Fallback to black if parsing fails.
                options.Color = Color.Black;
            }

            // Apply the text watermark.
            doc.Watermark.SetText(config.Text, options);

            // Save the watermarked document.
            string outputPath = Path.GetFileNameWithoutExtension(srcPath) + "_Watermarked.docx";
            doc.Save(outputPath);

            // Simple validation: ensure the output file exists.
            Console.WriteLine(File.Exists(outputPath)
                ? $"Watermark applied and saved to '{outputPath}'."
                : $"Failed to save watermarked document '{outputPath}'.");
        }
    }
}
