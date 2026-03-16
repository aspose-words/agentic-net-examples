using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace WatermarkBatchProcessor
{
    // Represents the JSON configuration for the watermark.
    public class WatermarkConfig
    {
        public string Text { get; set; }
        public string FontFamily { get; set; }
        public bool? IsSemitrasparent { get; set; }   // Controls opacity.
        public string ColorHex { get; set; }          // Optional color in hex format, e.g., "#FF0000".
        public float? FontSize { get; set; }          // Optional font size.
        public WatermarkLayout? Layout { get; set; }  // Optional layout (Diagonal or Horizontal).
    }

    public class Program
    {
        // Entry point.
        public static void Main(string[] args)
        {
            // Paths can be supplied via command‑line arguments or hard‑coded for testing.
            string configPath = args.Length > 0 ? args[0] : "watermarkConfig.json";
            string inputFolder = args.Length > 1 ? args[1] : "InputDocs";
            string outputFolder = args.Length > 2 ? args[2] : "OutputDocs";

            // Ensure output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Load watermark configuration.
            WatermarkConfig config = LoadConfig(configPath);

            // Gather all document files (doc, docx, dot, dotx, docm, dotm, odt, ott).
            string[] supportedExtensions = new[] { ".doc", ".docx", ".dot", ".dotx", ".docm", ".dotm", ".odt", ".ott" };
            List<string> inputFiles = new List<string>();
            foreach (string file in Directory.GetFiles(inputFolder))
            {
                if (Array.Exists(supportedExtensions, ext => ext.Equals(Path.GetExtension(file), StringComparison.OrdinalIgnoreCase)))
                {
                    inputFiles.Add(file);
                }
            }

            // Process each document.
            foreach (string inputPath in inputFiles)
            {
                // Load the document using the constructor that accepts a file name.
                Document doc = new Document(inputPath);

                // Prepare watermark options based on the configuration.
                TextWatermarkOptions options = new TextWatermarkOptions();

                if (!string.IsNullOrEmpty(config.FontFamily))
                    options.FontFamily = config.FontFamily;

                if (config.FontSize.HasValue)
                    options.FontSize = config.FontSize.Value;

                if (config.IsSemitrasparent.HasValue)
                    options.IsSemitrasparent = config.IsSemitrasparent.Value;

                if (config.Layout.HasValue)
                    options.Layout = config.Layout.Value;

                if (!string.IsNullOrEmpty(config.ColorHex))
                {
                    // Convert hex string to System.Drawing.Color.
                    Color parsedColor = ColorTranslator.FromHtml(config.ColorHex);
                    options.Color = parsedColor;
                }

                // Apply the text watermark using the SetText method with options.
                doc.Watermark.SetText(config.Text, options);

                // Determine output file path.
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(inputPath));

                // Save the modified document using the Save method that takes a file name.
                doc.Save(outputPath);
            }
        }

        // Reads and deserializes the JSON configuration file.
        private static WatermarkConfig LoadConfig(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException($"Configuration file not found: {path}");

            string json = File.ReadAllText(path);
            WatermarkConfig config = JsonSerializer.Deserialize<WatermarkConfig>(json);

            if (config == null)
                throw new InvalidOperationException("Failed to deserialize watermark configuration.");

            if (string.IsNullOrWhiteSpace(config.Text))
                throw new ArgumentException("Watermark text must be provided in the configuration.");

            return config;
        }
    }
}
