using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ConvertJpegToWebp
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a deterministic JPEG sample image using Aspose.Drawing.
        // -----------------------------------------------------------------
        const string jpegPath = "sample.jpg";
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        graphics.DrawRectangle(new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5), 20, 20, 160, 160);
        graphics.DrawString(
            "JPEG",
            new Aspose.Drawing.Font("Arial", 24),
            new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black),
            50,
            80);
        bitmap.Save(jpegPath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        graphics.Dispose();
        bitmap.Dispose();

        // ---------------------------------------------------------------
        // 2. Create a Word document and insert the JPEG image into it.
        // ---------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        const string docPath = "DocumentWithJpeg.docx";
        doc.Save(docPath);

        // ---------------------------------------------------------------
        // 3. Reload the document (demonstrates load/save lifecycle).
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // ---------------------------------------------------------------
        // 4. Extract JPEG images and convert each to high‑quality WebP.
        //    The conversion is performed by creating a temporary document
        //    that contains the extracted JPEG and saving that document as
        //    WebP using ImageSaveOptions (SaveFormat.WebP).
        // ---------------------------------------------------------------
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage) continue;
            if (shape.ImageData.ImageType != ImageType.Jpeg) continue;

            // Save the original JPEG image to a memory stream.
            using (MemoryStream jpegStream = new MemoryStream())
            {
                shape.ImageData.Save(jpegStream);
                jpegStream.Position = 0; // Reset before reuse.

                // Create a temporary document that contains only this image.
                Document tempDoc = new Document();
                DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
                // Insert the image from the memory stream.
                tempBuilder.InsertImage(jpegStream);

                // Prepare WebP output path.
                string webpPath = $"image_{imageIndex}.webp";

                // Save the temporary document as a WebP image (first page only).
                ImageSaveOptions webpOptions = new ImageSaveOptions(SaveFormat.WebP);
                tempDoc.Save(webpPath, webpOptions);

                // Validate that the WebP file was created.
                if (!File.Exists(webpPath))
                    throw new InvalidOperationException($"Failed to create WebP file: {webpPath}");

                Console.WriteLine($"Converted JPEG image #{imageIndex} to WebP: {webpPath}");
            }

            imageIndex++;
        }

        // Ensure at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found to convert.");

        // ---------------------------------------------------------------
        // 5. Cleanup sample files (optional).
        // ---------------------------------------------------------------
        File.Delete(jpegPath);
        File.Delete(docPath);
    }
}
