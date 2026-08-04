using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Directories for input BMPs and output WebP files.
        string inputDir = "InputImages";
        string outputDir = "OutputWebP";

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create deterministic BMP sample images.
        CreateSampleBmp(Path.Combine(inputDir, "red.bmp"), 100, 100, Aspose.Drawing.Color.Red);
        CreateSampleBmp(Path.Combine(inputDir, "green.bmp"), 100, 100, Aspose.Drawing.Color.Green);
        CreateSampleBmp(Path.Combine(inputDir, "blue.bmp"), 100, 100, Aspose.Drawing.Color.Blue);

        // Build a source document that contains the BMP images.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        foreach (string bmpPath in Directory.GetFiles(inputDir, "*.bmp"))
        {
            builder.InsertParagraph();
            // Insert the BMP image; Aspose.Words keeps the original format.
            builder.InsertImage(bmpPath);
        }

        string sourceDocPath = "SourceDocument.docx";
        sourceDoc.Save(sourceDocPath);

        // Load the document and extract image shapes.
        Document loadedDoc = new Document(sourceDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        var conversionLog = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Ensure the shape actually contains an image.
            if (!shape.HasImage)
                continue;

            // The sample images are BMPs, but to be safe we process all image shapes.
            // Save the original BMP size (if the image is BMP; otherwise size of the stored image).
            long originalSize = shape.ImageData.ImageBytes?.LongLength ?? 0;

            // Export the image to a memory stream.
            using (MemoryStream bmpStream = new MemoryStream())
            {
                shape.ImageData.Save(bmpStream);
                bmpStream.Position = 0;

                // Create a temporary document that holds the image.
                Document tempDoc = new Document();
                DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
                tempBuilder.InsertImage(bmpStream);

                // Define the output WebP file name.
                string webpPath = Path.Combine(outputDir, $"image_{imageIndex}.webp");

                // Save the temporary document as a WebP image (lossless by default).
                ImageSaveOptions webpOptions = new ImageSaveOptions(SaveFormat.WebP);
                tempDoc.Save(webpPath, webpOptions);

                // Log conversion details.
                long webpSize = new FileInfo(webpPath).Length;
                conversionLog.Add(
                    $"Image {imageIndex}: original size = {originalSize} bytes, WebP size = {webpSize} bytes, saved to '{webpPath}'.");
                imageIndex++;
            }
        }

        // Output the conversion log.
        Console.WriteLine("Batch BMP to WebP conversion completed.");
        foreach (string logEntry in conversionLog)
        {
            Console.WriteLine(logEntry);
        }

        // Validate that at least one image was converted.
        if (conversionLog.Count == 0)
            throw new InvalidOperationException("No images were found for conversion.");
    }

    // Helper method to create a deterministic BMP image.
    private static void CreateSampleBmp(string filePath, int width, int height, Aspose.Drawing.Color fillColor)
    {
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(fillColor);
        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();
    }
}
