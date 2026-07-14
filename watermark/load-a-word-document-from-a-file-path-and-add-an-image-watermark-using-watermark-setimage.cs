using System;
using System.IO;
using Aspose.Words;

namespace WatermarkExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare directories and file paths.
            string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
            Directory.CreateDirectory(dataDir);

            string docPath = Path.Combine(dataDir, "Sample.docx");
            string imagePath = Path.Combine(dataDir, "Watermark.png");
            string outputPath = Path.Combine(dataDir, "SampleWithWatermark.docx");

            // Create a simple Word document if it does not exist.
            if (!File.Exists(docPath))
            {
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln("This is a sample document.");
                doc.Save(docPath);
            }

            // Create a minimal 1x1 PNG image to use as a watermark if it does not exist.
            if (!File.Exists(imagePath))
            {
                // PNG byte array for a 1x1 transparent pixel.
                byte[] pngBytes = new byte[]
                {
                    0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                    0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                    0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                    0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
                    0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
                    0x54,0x78,0x9C,0x63,0x00,0x01,0x00,0x00,
                    0x05,0x00,0x01,0x0D,0x0A,0x2D,0xB4,0x00,
                    0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
                    0x42,0x60,0x82
                };
                File.WriteAllBytes(imagePath, pngBytes);
            }

            // Load the existing document.
            Document loadedDoc = new Document(docPath);

            // Add the image watermark using the Document.Watermark API.
            // Pass null for ImageWatermarkOptions to use defaults.
            loadedDoc.Watermark.SetImage(imagePath, null);

            // Save the document with the watermark applied.
            loadedDoc.Save(outputPath);

            // Indicate completion.
            Console.WriteLine("Watermark added and document saved to: " + outputPath);
        }
    }
}
