using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string inputFolder = "input";
        string outputFolder = "output";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a sample animated GIF (static for simplicity) if none exist
        string sampleGifPath = Path.Combine(inputFolder, "sample.gif");
        if (!File.Exists(sampleGifPath))
        {
            // Create a simple bitmap
            using (Bitmap bitmap = new Bitmap(200, 200))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.Clear(Aspose.Drawing.Color.White);
                    g.FillRectangle(new SolidBrush(Aspose.Drawing.Color.Blue), 20, 20, 160, 160);
                }
                // Save as GIF
                bitmap.Save(sampleGifPath, ImageFormat.Gif);
            }
        }

        // Process each GIF in the input folder
        string[] gifFiles = Directory.GetFiles(inputFolder, "*.gif");
        int processedCount = 0;

        foreach (string gifPath in gifFiles)
        {
            // Load the GIF image
            using (Image gifImage = Image.FromFile(gifPath))
            {
                // Determine output WebP path
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(gifPath);
                string webpPath = Path.Combine(outputFolder, fileNameWithoutExt + ".webp");

                // Save as WebP while preserving animation (if supported)
                // The Image.Save method infers format from file extension.
                gifImage.Save(webpPath);
                processedCount++;
            }
        }

        // Validate that at least one WebP file was created
        if (processedCount == 0)
            throw new InvalidOperationException("No GIF files were found for conversion.");

        string[] webpFiles = Directory.GetFiles(outputFolder, "*.webp");
        if (webpFiles.Length == 0)
            throw new InvalidOperationException("WebP conversion failed; no output files were created.");

        // Optionally, demonstrate loading a WebP back into a Word document using Aspose.Words
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string webpFile in webpFiles)
        {
            builder.InsertImage(webpFile);
        }
        string docPath = "ConvertedImages.docx";
        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new InvalidOperationException("Failed to save the Word document with WebP images.");
    }
}
