using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace WatermarkExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // (Optional) Add some sample text to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document with a text watermark.");

            // Configure custom font settings for the watermark.
            TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 36,
                Color = Color.Blue,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = false
            };

            // Add the text watermark using the configured options.
            doc.Watermark.SetText("Confidential", watermarkOptions);

            // Save the document to a file in the current directory.
            string outputPath = "WatermarkedDocument.docx";
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (System.IO.File.Exists(outputPath))
            {
                Console.WriteLine($"Watermarked document saved successfully to '{outputPath}'.");
            }
            else
            {
                Console.WriteLine("Failed to save the watermarked document.");
            }
        }
    }
}
