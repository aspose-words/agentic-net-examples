using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string inputBmpPath = "sample.bmp";
        const string docPath = "docWithBmp.docx";

        // 1. Create a sample BMP image (800x600) using Aspose.Drawing
        using (Bitmap bmp = new Bitmap(800, 600))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
            }

            bmp.Save(inputBmpPath, ImageFormat.Bmp);
        }

        // 2. Create a Word document and insert the BMP image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputBmpPath);
        doc.Save(docPath);

        // 3. Load the document and extract images
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Ensure the shape actually contains image data
            if (!shape.HasImage) continue;

            // 4. Save the original image bytes to a memory stream
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // reset before reading

                // 5. Load the image with Aspose.Drawing.Bitmap
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    int originalWidth = originalBitmap.Width;
                    int originalHeight = originalBitmap.Height;

                    // 6. Calculate new dimensions (width = 1024, height proportional)
                    const int targetWidth = 1024;
                    int targetHeight = (int)Math.Round((double)originalHeight * targetWidth / originalWidth);

                    // 7. Resize the bitmap
                    using (Bitmap resizedBitmap = new Bitmap(targetWidth, targetHeight))
                    {
                        using (Graphics g = Graphics.FromImage(resizedBitmap))
                        {
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.DrawImage(originalBitmap, 0, 0, targetWidth, targetHeight);
                        }

                        // 8. Save the resized BMP image
                        string resizedPath = $"resized_{imageIndex}.bmp";
                        resizedBitmap.Save(resizedPath, ImageFormat.Bmp);
                        imageIndex++;
                    }
                }
            }
        }

        // 9. Validation: ensure at least one resized image was created
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted and resized.");

        // Cleanup temporary files (optional)
        if (File.Exists(inputBmpPath)) File.Delete(inputBmpPath);
        if (File.Exists(docPath)) File.Delete(docPath);
    }
}
