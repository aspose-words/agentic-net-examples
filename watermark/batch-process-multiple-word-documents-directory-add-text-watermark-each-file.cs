using System;
using System.IO;
using System.Drawing;               // For Color
using Aspose.Words;                // Core document API
using Aspose.Words.Drawing;        // For WatermarkLayout

namespace WatermarkBatchProcessor
{
    class Program
    {
        static void Main()
        {
            // Folder containing the source Word documents.
            string sourceFolder = @"C:\Docs\Input";

            // Folder where watermarked documents will be saved.
            string targetFolder = @"C:\Docs\Output";

            // Ensure the target folder exists.
            Directory.CreateDirectory(targetFolder);

            // Text to be used as the watermark.
            const string watermarkText = "Confidential";

            // Iterate over all .docx files in the source folder.
            foreach (string sourcePath in Directory.GetFiles(sourceFolder, "*.docx"))
            {
                // Load the document from the file system.
                Document doc = new Document(sourcePath);

                // Configure watermark appearance.
                TextWatermarkOptions options = new TextWatermarkOptions
                {
                    FontFamily = "Arial",
                    FontSize = 36,
                    Color = Color.Red,
                    Layout = WatermarkLayout.Diagonal,
                    IsSemitrasparent = false
                };

                // Apply the text watermark to every page of the document.
                doc.Watermark.SetText(watermarkText, options);

                // Build the output file path (preserve original file name).
                string targetPath = Path.Combine(targetFolder, Path.GetFileName(sourcePath));

                // Save the modified document, overwriting if it already exists.
                doc.Save(targetPath);
            }

            Console.WriteLine("Watermarking completed.");
        }
    }
}
