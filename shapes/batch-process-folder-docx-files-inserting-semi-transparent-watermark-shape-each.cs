using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Determine base directory (the folder where the executable runs)
        string baseDir = AppContext.BaseDirectory;

        // Folder containing the source DOCX files (create if it doesn't exist)
        string sourceFolder = Path.Combine(baseDir, "Input");
        Directory.CreateDirectory(sourceFolder);

        // Folder where the processed files will be saved (create if it doesn't exist)
        string destinationFolder = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(destinationFolder);

        // Get all DOCX files in the source folder
        string[] docxFiles = Directory.GetFiles(sourceFolder, "*.docx", SearchOption.TopDirectoryOnly);

        if (docxFiles.Length == 0)
        {
            Console.WriteLine($"No DOCX files found in '{sourceFolder}'. Place files there and rerun the program.");
            return;
        }

        // Process each DOCX file
        foreach (string filePath in docxFiles)
        {
            // Load the document
            Document doc = new Document(filePath);

            // Configure watermark appearance
            TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 48,
                // Semi‑transparent gray using alpha channel (128 out of 255)
                Color = Color.FromArgb(128, Color.Gray),
                Layout = WatermarkLayout.Diagonal
            };

            // Apply the text watermark to the document
            doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

            // Save the modified document to the output folder
            string fileName = Path.GetFileName(filePath);
            string outputPath = Path.Combine(destinationFolder, fileName);
            doc.Save(outputPath);
        }

        Console.WriteLine($"Processed {docxFiles.Length} file(s). Output saved to '{destinationFolder}'.");
    }
}
