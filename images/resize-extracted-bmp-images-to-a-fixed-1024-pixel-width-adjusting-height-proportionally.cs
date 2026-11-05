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
        // Create a deterministic sample BMP image.
        const string sampleBmpPath = "sample.bmp";
        CreateSampleBmp(sampleBmpPath, 200, 200);

        // Build a Word document and insert the BMP image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleBmpPath);
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Reload the document to simulate a fresh processing scenario.
        Document loadedDoc = new Document(docPath);

        // Extract all Shape nodes that contain images, resize them to 1024px width,
        // keep the aspect ratio, and save the resized images to disk.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int resizedCount = 0;

        for (int i = 0; i < shapeNodes.Count; i++)
        {
            Shape shape = (Shape)shapeNodes[i];
            if (!shape.HasImage)
                continue; // Skip non‑image shapes.

            // Obtain the original image bytes.
            byte[] originalBytes = shape.ImageData.ImageBytes;
            using (MemoryStream originalStream = new MemoryStream(originalBytes))
            {
                originalStream.Position = 0;
                using (Bitmap originalBmp = new Bitmap(originalStream))
                {
                    int originalWidth = originalBmp.Width;
                    int originalHeight = originalBmp.Height;

                    const int targetWidth = 1024;
                    int targetHeight = (int)Math.Round((double)originalHeight * targetWidth / originalWidth);

                    // Create a new bitmap with the target dimensions.
                    using (Bitmap resizedBmp = new Bitmap(targetWidth, targetHeight))
                    {
                        using (Graphics g = Graphics.FromImage(resizedBmp))
                        {
                            g.Clear(Aspose.Drawing.Color.White);
                            g.DrawImage(originalBmp, new Rectangle(0, 0, targetWidth, targetHeight));
                        }

                        // Save the resized bitmap to a deterministic file name.
                        string resizedPath = $"resized-{i + 1}.bmp";
                        resizedBmp.Save(resizedPath, ImageFormat.Bmp);

                        // Replace the image inside the document with the resized version.
                        shape.ImageData.SetImage(resizedPath);

                        resizedCount++;
                    }
                }
            }
        }

        // Validate that at least one image was resized.
        if (resizedCount == 0)
            throw new InvalidOperationException("No images were found and resized.");

        // Save the document with the resized images.
        const string updatedDocPath = "sample_resized.docx";
        loadedDoc.Save(updatedDocPath);
    }

    // Helper method to create a simple BMP image using Aspose.Drawing.
    private static void CreateSampleBmp(string path, int width, int height)
    {
        using (Bitmap bmp = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Pen pen = new Pen(Aspose.Drawing.Color.Black, 5))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bmp.Save(path, ImageFormat.Bmp);
        }
    }
}
