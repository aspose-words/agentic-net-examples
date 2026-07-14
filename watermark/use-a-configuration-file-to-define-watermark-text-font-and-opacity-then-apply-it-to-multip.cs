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
        public string Text { get; set; } = "Sample Watermark";
        public string FontFamily { get; set; } = "Arial";
        public float FontSize { get; set; } = 36;
        public string ColorHex { get; set; } = "#FF0000"; // Red
        public bool IsSemitrasparent { get; set; } = false;
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Ensure a configuration file exists. If not, create a default one.
        // -----------------------------------------------------------------
        const string configFile = "watermarkConfig.json";
        if (!File.Exists(configFile))
        {
            var defaultConfig = new WatermarkConfig();
            string json = JsonSerializer.Serialize(defaultConfig, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(configFile, json);
        }

        // -----------------------------------------------------------------
        // 2. Load the configuration.
        // -----------------------------------------------------------------
        WatermarkConfig config = JsonSerializer.Deserialize<WatermarkConfig>(File.ReadAllText(configFile))!;

        // Convert the hex color string to a System.Drawing.Color.
        Color watermarkColor = ColorTranslator.FromHtml(config.ColorHex);

        // -----------------------------------------------------------------
        // 3. Prepare the watermark options based on the configuration.
        // -----------------------------------------------------------------
        var textOptions = new TextWatermarkOptions
        {
            FontFamily = config.FontFamily,
            FontSize = config.FontSize,
            Color = watermarkColor,
            IsSemitrasparent = config.IsSemitrasparent,
            Layout = WatermarkLayout.Diagonal
        };

        // -----------------------------------------------------------------
        // 4. Create an output folder for the generated documents.
        // -----------------------------------------------------------------
        const string outputFolder = "Output";
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // 5. Process multiple documents, applying the same watermark.
        // -----------------------------------------------------------------
        string[] sourceNames = { "Doc1.docx", "Doc2.docx", "Doc3.docx" };
        foreach (string sourceName in sourceNames)
        {
            // Create a new blank document.
            var doc = new Document();

            // Add a simple paragraph so the document is not empty.
            var builder = new DocumentBuilder(doc);
            builder.Writeln($"This is the content of {Path.GetFileNameWithoutExtension(sourceName)}.");

            // Apply the text watermark using the configuration.
            doc.Watermark.SetText(config.Text, textOptions);

            // Save the watermarked document.
            string outputPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(sourceName) + "_Watermarked.docx");
            doc.Save(outputPath);
        }

        // -----------------------------------------------------------------
        // 6. Simple validation: ensure that output files were created.
        // -----------------------------------------------------------------
        foreach (string sourceName in sourceNames)
        {
            string outputPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(sourceName) + "_Watermarked.docx");
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Created: {outputPath}");
            }
        }
    }
}
