using System;
using System.IO;
using Aspose.Words;

public class Program
{
    // Entry point of the console application.
    public static void Main(string[] args)
    {
        // Expect two arguments: directory path and watermark text.
        if (args.Length < 2)
        {
            // No interactive prompts; simply exit if arguments are insufficient.
            return;
        }

        string directoryPath = args[0];
        string watermarkText = args[1];

        // Ensure the target directory exists. If it does not, create it and add a sample document.
        if (!Directory.Exists(directoryPath))
        {
            Directory.CreateDirectory(directoryPath);
            CreateSampleDocument(Path.Combine(directoryPath, "Sample.docx"));
        }

        // Process each .docx file in the directory.
        foreach (string filePath in Directory.GetFiles(directoryPath, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Apply a text watermark using the native API.
            doc.Watermark.SetText(watermarkText);

            // Overwrite the original file with the watermarked version.
            doc.Save(filePath);
        }
    }

    // Creates a simple Word document with placeholder content.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document generated for watermark processing.");
        doc.Save(filePath);
    }
}
