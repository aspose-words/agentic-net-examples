using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic BMP image larger than 1024px width.
        const string inputImagePath = "input.bmp";
        const int originalWidth = 2000;
        const int originalHeight = 1500;
        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Simple visual content.
            graphics.DrawRectangle(Pens.Black, 10, 10, originalWidth - 20, originalHeight - 20);
            bitmap.Save(inputImagePath);
        }

        // Insert the BMP image into a Word document.
        const string docPath = "original.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // Load the document and extract images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the image data to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading.

                // Load the original bitmap.
                using (Bitmap originalBmp = new Bitmap(imageStream))
                {
                    // Compute new dimensions preserving aspect ratio.
                    const int targetWidth = 1024;
                    int newHeight = (int)Math.Round((double)originalBmp.Height * targetWidth / originalBmp.Width);

                    // Resize the bitmap.
                    using (Bitmap resizedBmp = new Bitmap(targetWidth, newHeight))
                    using (Graphics g = Graphics.FromImage(resizedBmp))
                    {
                        g.DrawImage(originalBmp, 0, 0, targetWidth, newHeight);

                        // Save the resized image to a deterministic file.
                        string outputPath = $"resized_{extractedCount}.bmp";
                        resizedBmp.Save(outputPath);

                        // Validate that the file was created.
                        if (!File.Exists(outputPath))
                            throw new InvalidOperationException($"Failed to create resized image file: {outputPath}");
                    }
                }

                extractedCount++;
            }
        }

        // Ensure at least one image was processed.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Cleanup temporary files (optional).
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
    }
}
