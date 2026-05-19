using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample JPEG image using Aspose.Drawing.
        string jpegPath = Path.Combine(artifactsDir, "sample.jpg");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                g.FillRectangle(new SolidBrush(Aspose.Drawing.Color.Red), 0, 0, 200, 200);
            }
            bitmap.Save(jpegPath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }

        // 2. Insert the JPEG into a Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        string docPath = Path.Combine(artifactsDir, "doc_with_image.docx");
        doc.Save(docPath);

        // 3. Extract JPEG images from the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Save the image data to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                // 4. Create a temporary document that contains only this image.
                Document tempDoc = new Document();
                DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
                tempBuilder.InsertImage(imageStream);

                // 5. Convert the image to high‑quality WebP.
                ImageSaveOptions webpOptions = new ImageSaveOptions(SaveFormat.WebP)
                {
                    // High quality for WebP (the property affects JPEG quality; for WebP we keep defaults).
                    JpegQuality = 100
                };

                string webpPath = Path.Combine(artifactsDir, $"image_{imageIndex}.webp");
                tempDoc.Save(webpPath, webpOptions);

                // Validate that the WebP file was created.
                if (!File.Exists(webpPath))
                    throw new InvalidOperationException($"Failed to create WebP file: {webpPath}");

                imageIndex++;
            }
        }

        // Ensure at least one image was converted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found to convert.");

        // Example completed successfully.
        Console.WriteLine($"Converted {imageIndex} JPEG image(s) to WebP. Files are located in: {artifactsDir}");
    }
}
