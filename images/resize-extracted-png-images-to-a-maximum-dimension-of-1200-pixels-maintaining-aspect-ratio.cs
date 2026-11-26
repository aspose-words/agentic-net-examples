using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample PNG image (2000x1500) using Aspose.Drawing.
        const string inputImagePath = "input.png";
        const int sampleWidth = 2000;
        const int sampleHeight = 1500;
        using (Bitmap bitmap = new Bitmap(sampleWidth, sampleHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
            }
            bitmap.Save(inputImagePath);
        }

        // Create a Word document and insert the sample image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // Load the document and extract PNG images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the extracted original PNG.
            string extractedPath = $"extracted-{imageIndex}.png";
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0;
                using (FileStream fs = new FileStream(extractedPath, FileMode.Create, FileAccess.Write))
                {
                    ms.CopyTo(fs);
                }

                // Load the image for resizing.
                ms.Position = 0;
                using (Bitmap originalBitmap = new Bitmap(ms))
                {
                    // Determine new dimensions while maintaining aspect ratio.
                    const int maxDimension = 1200;
                    double widthScale = (double)maxDimension / originalBitmap.Width;
                    double heightScale = (double)maxDimension / originalBitmap.Height;
                    double scale = Math.Min(widthScale, heightScale);
                    if (scale > 1.0) // Image already within limits; keep original size.
                        scale = 1.0;

                    int newWidth = (int)(originalBitmap.Width * scale);
                    int newHeight = (int)(originalBitmap.Height * scale);

                    // Resize the image.
                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics g = Graphics.FromImage(resizedBitmap))
                        {
                            g.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                        }

                        // Save the resized PNG.
                        string resizedPath = $"resized-{imageIndex}.png";
                        resizedBitmap.Save(resizedPath);
                    }
                }
            }

            imageIndex++;
        }

        // Validate that at least one resized image was created.
        string[] resizedFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "resized-*.png");
        if (resizedFiles.Length == 0)
            throw new InvalidOperationException("No resized PNG images were produced.");

        // Cleanup: optional (commented out if you want to keep files).
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
    }
}
