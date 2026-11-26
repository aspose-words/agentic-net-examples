using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a deterministic BMP image (200x200, solid red)
        // ------------------------------------------------------------
        const string sampleBmpPath = "sample.bmp";
        using (Bitmap bmp = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Aspose.Drawing.Color.Red);
            }
            bmp.Save(sampleBmpPath, ImageFormat.Bmp);
        }

        // ------------------------------------------------------------
        // 2. Insert the BMP image into a new Word document
        // ------------------------------------------------------------
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleBmpPath);
        doc.Save(docPath);

        // ------------------------------------------------------------
        // 3. Load the document and extract all image shapes
        // ------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int resizedCount = 0;
        int imageIndex = 0;
        const int targetWidth = 1024;

        foreach (Shape shape in shapeNodes)
        {
            // Ensure the shape actually contains an image.
            if (!shape.HasImage)
                continue;

            // ----------------------------------------------------------------
            // Extract the original image bytes and load them into a bitmap.
            // ----------------------------------------------------------------
            byte[] imageBytes = shape.ImageData.ImageBytes;
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0; // reset just in case
                using (Bitmap originalBitmap = new Bitmap(ms))
                {
                    int originalWidth = originalBitmap.Width;
                    int originalHeight = originalBitmap.Height;

                    // Determine new dimensions: width = 1024 (or original if smaller),
                    // height scaled proportionally.
                    int newWidth = originalWidth >= targetWidth ? targetWidth : originalWidth;
                    int newHeight = (int)(originalHeight * (newWidth / (double)originalWidth));

                    // If dimensions are unchanged, skip resizing but still count as processed.
                    if (newWidth == originalWidth && newHeight == originalHeight)
                    {
                        // Save a copy of the original image to keep the output pattern consistent.
                        string unchangedPath = $"resized-{imageIndex}.bmp";
                        originalBitmap.Save(unchangedPath, ImageFormat.Bmp);
                        if (!File.Exists(unchangedPath))
                            throw new InvalidOperationException($"Failed to save image: {unchangedPath}");

                        resizedCount++;
                        imageIndex++;
                        continue;
                    }

                    // --------------------------------------------------------
                    // Create the resized bitmap and draw the scaled image.
                    // --------------------------------------------------------
                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics g = Graphics.FromImage(resizedBitmap))
                        {
                            g.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                        }

                        // Save the resized bitmap as BMP.
                        string resizedPath = $"resized-{imageIndex}.bmp";
                        resizedBitmap.Save(resizedPath, ImageFormat.Bmp);
                        if (!File.Exists(resizedPath))
                            throw new InvalidOperationException($"Failed to save resized image: {resizedPath}");

                        resizedCount++;
                        imageIndex++;
                    }
                }
            }
        }

        // ------------------------------------------------------------
        // Validation: at least one image should have been processed.
        // ------------------------------------------------------------
        if (resizedCount == 0)
            throw new InvalidOperationException("No images were found and processed in the document.");

        // Optional cleanup of temporary files (commented out).
        // File.Delete(sampleBmpPath);
        // File.Delete(docPath);
    }
}
