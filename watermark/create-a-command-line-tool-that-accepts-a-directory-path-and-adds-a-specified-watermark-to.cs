using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main(string[] args)
    {
        // Expect at least the directory path argument.
        if (args.Length == 0)
            return;

        string directoryPath = args[0];
        string watermarkText = args.Length > 1 ? args[1] : "Aspose Watermark";

        // Ensure the directory exists.
        if (!Directory.Exists(directoryPath))
            Directory.CreateDirectory(directoryPath);

        // If the directory contains no Word files, create a sample document.
        string[] wordFiles = Directory.GetFiles(directoryPath, "*.docx");
        if (wordFiles.Length == 0)
        {
            string samplePath = Path.Combine(directoryPath, "Sample.docx");
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("This is a sample document.");
            sampleDoc.Save(samplePath);
            wordFiles = new[] { samplePath };
        }

        // Process each .docx file: add the text watermark and overwrite the file.
        foreach (string filePath in wordFiles)
        {
            Document doc = new Document(filePath);
            doc.Watermark.SetText(watermarkText);
            doc.Save(filePath);
        }
    }
}
