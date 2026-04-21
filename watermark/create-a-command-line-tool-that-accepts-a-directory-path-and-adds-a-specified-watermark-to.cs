using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class WatermarkTool
{
    public static void Main(string[] args)
    {
        // Expect two arguments: directory path and watermark text.
        if (args.Length < 2)
        {
            // Insufficient arguments; exit silently.
            return;
        }

        string directoryPath = args[0];
        string watermarkText = args[1];

        // Ensure the directory exists; if not, create it and a sample document.
        if (!Directory.Exists(directoryPath))
        {
            Directory.CreateDirectory(directoryPath);
            CreateSampleDocument(Path.Combine(directoryPath, "Sample.docx"));
        }

        // Process all .docx files in the directory.
        string[] docFiles = Directory.GetFiles(directoryPath, "*.docx", SearchOption.TopDirectoryOnly);
        foreach (string filePath in docFiles)
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Apply a text watermark.
            doc.Watermark.SetText(watermarkText);

            // Overwrite the original file with the watermarked version.
            doc.Save(filePath);
        }
    }

    // Creates a simple sample Word document with placeholder text.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        doc.Save(filePath);
    }
}
