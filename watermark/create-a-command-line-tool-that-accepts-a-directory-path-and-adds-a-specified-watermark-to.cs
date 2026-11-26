using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main(string[] args)
    {
        // Determine the target directory and watermark text.
        string targetDirectory;
        string watermarkText;

        if (args.Length >= 2)
        {
            targetDirectory = args[0];
            watermarkText = args[1];
        }
        else
        {
            // Fallback: create a temporary directory with a sample document.
            targetDirectory = Path.Combine(Path.GetTempPath(), "AsposeWatermarkDemo");
            watermarkText = "Sample Watermark";

            if (!Directory.Exists(targetDirectory))
                Directory.CreateDirectory(targetDirectory);

            // Create a simple sample document if none exist.
            string samplePath = Path.Combine(targetDirectory, "Sample.docx");
            if (!File.Exists(samplePath))
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln("This is a sample document.");
                sampleDoc.Save(samplePath);
            }
        }

        // Ensure the directory exists.
        if (!Directory.Exists(targetDirectory))
            Directory.CreateDirectory(targetDirectory);

        // Process each .docx file in the directory.
        string[] docFiles = Directory.GetFiles(targetDirectory, "*.docx", SearchOption.TopDirectoryOnly);
        foreach (string filePath in docFiles)
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Apply the text watermark.
            doc.Watermark.SetText(watermarkText);

            // Save the document, overwriting the original file.
            doc.Save(filePath);
        }

        // Optionally, indicate completion (no interactive prompts).
        Console.WriteLine($"Processed {docFiles.Length} document(s) in \"{targetDirectory}\".");
    }
}
