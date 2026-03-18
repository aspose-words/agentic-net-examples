using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Folder containing the source DOCX files
        string sourceFolder = @"C:\Docs\Input";

        // Folder where the processed files will be saved
        string destinationFolder = @"C:\Docs\Output";

        // Ensure the destination folder exists
        Directory.CreateDirectory(destinationFolder);

        // Process each DOCX file in the source folder
        foreach (string filePath in Directory.GetFiles(sourceFolder, "*.docx"))
        {
            // Load the document (uses the Document(string) constructor)
            Document doc = new Document(filePath);

            // Configure watermark appearance
            TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 48,
                Color = System.Drawing.Color.Gray,          // Gray color for the watermark
                Layout = WatermarkLayout.Diagonal,         // Diagonal layout
                IsSemitrasparent = true                    // Semi‑transparent
            };

            // Apply the text watermark to the document
            doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

            // Save the modified document (overwrites or writes to a new location)
            string fileName = Path.GetFileName(filePath);
            string outputPath = Path.Combine(destinationFolder, fileName);
            doc.Save(outputPath);
        }
    }
}
