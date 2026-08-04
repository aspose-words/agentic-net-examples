using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class BatchImageConversion
{
    public static void Main()
    {
        // Define folders for sample documents and output images.
        string baseDir = Directory.GetCurrentDirectory();
        string docsDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "WebPImages");

        // Ensure clean environment.
        if (Directory.Exists(docsDir)) Directory.Delete(docsDir, true);
        if (Directory.Exists(outputDir)) Directory.Delete(outputDir, true);
        Directory.CreateDirectory(docsDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create sample images (PNG and JPEG) using Aspose.Drawing.
        // -----------------------------------------------------------------
        string pngPath = Path.Combine(baseDir, "sample.png");
        string jpegPath = Path.Combine(baseDir, "sample.jpg");

        CreateSampleImage(pngPath, 200, 200, Aspose.Drawing.Color.Blue);
        CreateSampleImage(jpegPath, 200, 200, Aspose.Drawing.Color.Green);

        // -----------------------------------------------------------------
        // 2. Create sample Word documents that contain the images.
        // -----------------------------------------------------------------
        for (int docIndex = 1; docIndex <= 2; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {docIndex}");
            builder.InsertImage(pngPath);
            builder.InsertParagraph();
            builder.InsertImage(jpegPath);

            string docPath = Path.Combine(docsDir, $"SampleDoc{docIndex}.docx");
            doc.Save(docPath);
        }

        // -----------------------------------------------------------------
        // 3. Batch process each Word file: extract images and convert to WebP.
        // -----------------------------------------------------------------
        int totalConverted = 0;
        foreach (string docFile in Directory.GetFiles(docsDir, "*.docx"))
        {
            Document srcDoc = new Document(docFile);
            NodeCollection shapeNodes = srcDoc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (!shape.HasImage) continue;

                // Extract image bytes.
                using (MemoryStream imgStream = new MemoryStream())
                {
                    shape.ImageData.Save(imgStream);
                    imgStream.Position = 0;

                    // Insert the extracted image into a temporary document.
                    Document tempDoc = new Document();
                    DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
                    tempBuilder.InsertImage(imgStream);

                    // Prepare WebP output path.
                    string webpFileName = $"{Path.GetFileNameWithoutExtension(docFile)}_img{imageIndex}.webp";
                    string webpPath = Path.Combine(outputDir, webpFileName);

                    // Save the temporary document as a WebP image.
                    ImageSaveOptions webpOptions = new ImageSaveOptions(SaveFormat.WebP);
                    tempDoc.Save(webpPath, webpOptions);

                    // Validate that the WebP file was created.
                    if (!File.Exists(webpPath))
                        throw new InvalidOperationException($"Failed to create WebP file: {webpPath}");

                    totalConverted++;
                }

                imageIndex++;
            }
        }

        // Ensure at least one image was converted.
        if (totalConverted == 0)
            throw new InvalidOperationException("No images were found and converted.");

        Console.WriteLine($"Converted {totalConverted} images to WebP format in folder: {outputDir}");
    }

    // Helper method to create a deterministic bitmap and save it to a file.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color fillColor)
    {
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(fillColor);
            }
            bitmap.Save(filePath);
        }
    }
}
