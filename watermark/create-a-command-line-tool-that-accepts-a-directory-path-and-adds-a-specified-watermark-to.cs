using System;
using System.IO;
using Aspose.Words;

public class WatermarkTool
{
    public static void Main(string[] args)
    {
        // Determine the target directory and watermark text from command‑line arguments.
        string targetDirectory = args.Length > 0 ? args[0] : "InputDocs";
        string watermarkText = args.Length > 1 ? args[1] : "Sample Watermark";

        // Ensure the directory exists.
        if (!Directory.Exists(targetDirectory))
        {
            Directory.CreateDirectory(targetDirectory);
        }

        // If the directory contains no Word documents, create a sample document.
        if (Directory.GetFiles(targetDirectory, "*.docx").Length == 0)
        {
            var sampleDoc = new Document();
            var builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("This is a sample document.");
            string samplePath = Path.Combine(targetDirectory, "Sample.docx");
            sampleDoc.Save(samplePath);
        }

        // Process each .docx file in the directory.
        foreach (string filePath in Directory.GetFiles(targetDirectory, "*.docx"))
        {
            // Load the document.
            var doc = new Document(filePath);

            // Apply a text watermark using the native API.
            doc.Watermark.SetText(watermarkText);

            // Overwrite the original file with the watermarked version.
            doc.Save(filePath);
        }

        // Optional: simple validation that files exist after processing.
        foreach (string filePath in Directory.GetFiles(targetDirectory, "*.docx"))
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Failed to save watermarked file: {filePath}");
            }
        }
    }
}
