using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a folder to hold sample documents.
        string docsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(docsFolder);

        // Generate a few sample Word documents.
        for (int i = 1; i <= 3; i++)
        {
            string samplePath = Path.Combine(docsFolder, $"Sample{i}.docx");
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is sample document {i}.");
            sampleDoc.Save(samplePath);
        }

        // Process each .docx file in the folder and add a text watermark.
        foreach (string filePath in Directory.GetFiles(docsFolder, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Add a text watermark.
            doc.Watermark.SetText("CONFIDENTIAL");

            // Save the watermarked document with a new name.
            string outputPath = Path.Combine(
                docsFolder,
                Path.GetFileNameWithoutExtension(filePath) + "_watermarked.docx");

            doc.Save(outputPath);

            // Simple validation that the output file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Watermark added to: {Path.GetFileName(outputPath)}");
            }
        }
    }
}
