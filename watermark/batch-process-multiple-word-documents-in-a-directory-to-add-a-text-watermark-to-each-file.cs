using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define a folder for sample input documents.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        Directory.CreateDirectory(inputFolder);

        // Create a few sample Word documents.
        for (int i = 1; i <= 3; i++)
        {
            string samplePath = Path.Combine(inputFolder, $"Sample{i}.docx");
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i}.");
            doc.Save(samplePath);
        }

        // Batch process each .docx file in the folder and add a text watermark.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            var doc = new Document(filePath);
            // Apply a simple text watermark.
            doc.Watermark.SetText("CONFIDENTIAL");
            // Overwrite the original file with the watermarked version.
            doc.Save(filePath);
        }

        // Indicate completion.
        Console.WriteLine("Text watermark applied to all documents.");
    }
}
