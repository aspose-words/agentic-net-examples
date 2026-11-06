using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample PNG image.
        const string sampleImagePath = "sample.png";
        const int imgWidth = 100;
        const int imgHeight = 100;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Color.White);
                // Draw a red rectangle.
                graphics.FillRectangle(new SolidBrush(Color.Red), 20, 20, 60, 60);
            }
            bitmap.Save(sampleImagePath);
        }

        // Step 2: Create a Word document and insert the sample image.
        const string docPath = "DocumentWithImage.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Step 3: Load the document and process each PNG image.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int invertedCount = 0;

        for (int i = 0; i < shapeNodes.Count; i++)
        {
            Shape shape = (Shape)shapeNodes[i];
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Extract the image to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading.

                // Load the image into a bitmap for manipulation.
                using (Bitmap bmp = new Bitmap(imageStream))
                {
                    // Invert colors pixel by pixel.
                    for (int y = 0; y < bmp.Height; y++)
                    {
                        for (int x = 0; x < bmp.Width; x++)
                        {
                            Color original = bmp.GetPixel(x, y);
                            Color inverted = Color.FromArgb(
                                255 - original.R,
                                255 - original.G,
                                255 - original.B);
                            bmp.SetPixel(x, y, inverted);
                        }
                    }

                    // Save the inverted image.
                    string outputPath = $"inverted-{i}.png";
                    bmp.Save(outputPath);
                    if (!File.Exists(outputPath))
                        throw new InvalidOperationException($"Failed to save inverted image: {outputPath}");

                    invertedCount++;
                }
            }
        }

        // Validation: ensure at least one image was processed.
        if (invertedCount == 0)
            throw new InvalidOperationException("No PNG images were found and processed in the document.");

        // Clean up temporary files (optional).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
