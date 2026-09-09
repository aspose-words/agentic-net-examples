using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ConvertJpegToWebP
{
    public static void Main()
    {
        // Prepare deterministic file names.
        const string jpegPath = "sample.jpg";
        const string docPath = "sample.docx";

        // -------------------------------------------------
        // 1. Create a sample JPEG image using Aspose.Drawing.
        // -------------------------------------------------
        int width = 200;
        int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.Red);
            }
            // Save as JPEG.
            bitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // -------------------------------------------------
        // 2. Insert the JPEG image into a Word document.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract JPEG images.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Save the extracted JPEG.
            string extractedJpeg = $"extracted_{imageIndex}.jpg";
            shape.ImageData.Save(extractedJpeg);

            // -------------------------------------------------
            // 4. Convert the extracted JPEG to high‑quality WebP.
            // -------------------------------------------------
            // Load the JPEG into a memory stream.
            using (MemoryStream jpegStream = new MemoryStream())
            {
                shape.ImageData.Save(jpegStream);
                jpegStream.Position = 0;

                // Create a temporary document that contains only this image.
                Document tempDoc = new Document();
                DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
                tempBuilder.InsertImage(jpegStream);

                // Save the document page as WebP.
                string webpPath = $"converted_{imageIndex}.webp";
                ImageSaveOptions webpOptions = new ImageSaveOptions(SaveFormat.WebP);
                // High quality: set the JPEG quality property (used for lossy formats).
                webpOptions.JpegQuality = 100;
                tempDoc.Save(webpPath, webpOptions);

                // Validate that the WebP file was created.
                if (!File.Exists(webpPath))
                    throw new InvalidOperationException($"WebP conversion failed for image index {imageIndex}.");
            }

            imageIndex++;
        }

        // Final validation: at least one conversion should have occurred.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found to convert.");

        // Clean up temporary files (optional).
        // File.Delete(jpegPath);
        // File.Delete(docPath);
    }
}
