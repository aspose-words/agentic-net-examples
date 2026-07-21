using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample BMP image.
        const string sampleBmpPath = "sample.bmp";
        using (Bitmap bmp = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.LightBlue);
            }
            bmp.Save(sampleBmpPath);
        }

        // Step 2: Insert the BMP image into a Word document.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleBmpPath);
        doc.Save(docPath);

        // Step 3: Reload the document and extract all images (including BMP).
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image using Aspose.Drawing.Bitmap.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Calculate new dimensions (fixed width 1024px, proportional height).
                    int newWidth = 1024;
                    int newHeight = (int)Math.Round((double)originalBitmap.Height * newWidth / originalBitmap.Width);

                    // Resize the bitmap.
                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                        {
                            graphics.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                        }

                        // Save the resized image as BMP.
                        string resizedPath = $"resized-{imageIndex}.bmp";
                        resizedBitmap.Save(resizedPath);
                        Console.WriteLine($"Resized image saved to: {resizedPath}");
                    }
                }
            }

            imageIndex++;
        }

        // Validation: ensure at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were found and resized.");

        // Optional cleanup (commented out to keep output files for verification).
        // File.Delete(sampleBmpPath);
        // File.Delete(docPath);
    }
}
