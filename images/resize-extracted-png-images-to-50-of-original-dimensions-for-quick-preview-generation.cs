using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string inputImagePath = "input.png";
        const string documentPath = "doc_with_image.docx";

        // ------------------------------------------------------------
        // 1. Create a sample PNG image (200x200) using Aspose.Drawing.
        // ------------------------------------------------------------
        int originalWidth = 200;
        int originalHeight = 200;
        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                g.Clear(Color.White);
                // Draw a simple red rectangle.
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    g.DrawRectangle(pen, 10, 10, originalWidth - 20, originalHeight - 20);
                }
            }
            // Save the image to a deterministic file.
            bitmap.Save(inputImagePath);
        }

        // ------------------------------------------------------------
        // 2. Create a Word document and insert the PNG image.
        // ------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(documentPath);

        // ------------------------------------------------------------
        // 3. Load the document (optional, we already have it) and extract PNG images.
        // ------------------------------------------------------------
        // Ensure we work with the same document instance.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int previewIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // --------------------------------------------------------
            // 4. Save the image data to a memory stream.
            // --------------------------------------------------------
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading.

                // ----------------------------------------------------
                // 5. Load the image into a Bitmap.
                // ----------------------------------------------------
                using (Bitmap originalBitmap = new Bitmap(imageStream))
                {
                    // Compute 50% dimensions.
                    int newWidth = originalBitmap.Width / 2;
                    int newHeight = originalBitmap.Height / 2;

                    // Guard against zero dimensions.
                    if (newWidth == 0) newWidth = 1;
                    if (newHeight == 0) newHeight = 1;

                    // ------------------------------------------------
                    // 6. Create a resized bitmap.
                    // ------------------------------------------------
                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics g = Graphics.FromImage(resizedBitmap))
                        {
                            // Draw the original image scaled to the new size.
                            g.DrawImage(
                                originalBitmap,
                                new Rectangle(0, 0, newWidth, newHeight));
                        }

                        // ------------------------------------------------
                        // 7. Save the preview image.
                        // ------------------------------------------------
                        string previewPath = $"preview_{previewIndex}.png";
                        resizedBitmap.Save(previewPath);
                        previewIndex++;
                    }
                }
            }
        }

        // ------------------------------------------------------------
        // 8. Validation: ensure at least one preview file was created.
        // ------------------------------------------------------------
        if (previewIndex == 0)
            throw new InvalidOperationException("No PNG images were extracted and resized.");

        // Optional: clean up the temporary files (comment out if inspection needed).
        // File.Delete(inputImagePath);
        // File.Delete(documentPath);
    }
}
