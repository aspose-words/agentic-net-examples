using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Define folders for input and output documents.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample source documents if the input folder is empty.
        if (Directory.GetFiles(inputFolder, "*.docx").Length == 0)
        {
            for (int i = 1; i <= 3; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"This is sample document {i}.");
                string samplePath = Path.Combine(inputFolder, $"Sample{i}.docx");
                sampleDoc.Save(samplePath);
            }
        }

        // Prepare watermark options (optional, can be omitted for default settings).
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };

        // Process each .docx file in the input folder.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Add a text watermark.
            doc.Watermark.SetText("Confidential", watermarkOptions);

            // Save the watermarked document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }

        // Optional: indicate completion (no interactive input required).
        Console.WriteLine("Batch watermarking completed.");
    }
}
