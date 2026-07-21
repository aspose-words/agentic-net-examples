using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main(string[] args)
    {
        // Determine input directory and watermark text from command‑line arguments.
        string inputDir = args.Length > 0 ? args[0] : "InputDocs";
        string watermarkText = args.Length > 1 ? args[1] : "Sample Watermark";

        // Ensure the directory exists.
        Directory.CreateDirectory(inputDir);

        // If the directory is empty, create a sample document to demonstrate the watermarking.
        if (Directory.GetFiles(inputDir, "*.docx").Length == 0)
        {
            Document sampleDoc = new Document();
            var builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("This is a sample document.");
            string samplePath = Path.Combine(inputDir, "Sample.docx");
            sampleDoc.Save(samplePath);
        }

        // Process each .docx file in the directory.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Apply the text watermark.
            doc.Watermark.SetText(watermarkText);

            // Overwrite the original file with the watermarked version.
            doc.Save(filePath);
        }

        // Optionally, indicate completion (no interactive input required).
        Console.WriteLine("Watermarking completed.");
    }
}
