using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image (input.png)
        const int sampleWidth = 200;
        const int sampleHeight = 300;
        string inputImagePath = "input.png";

        using (Bitmap bitmap = new Bitmap(sampleWidth, sampleHeight))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);
            // Draw a simple rectangle to make the image recognizable
            g.FillRectangle(new SolidBrush(Color.LightBlue), 20, 20, sampleWidth - 40, sampleHeight - 40);
            bitmap.Save(inputImagePath);
        }

        // Create a Word document and insert the sample image
        string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // Load the document for image extraction
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        int index = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Extract the image to a memory stream
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                // Load the extracted image into a bitmap
                using (Bitmap originalBitmap = new Bitmap(imageStream))
                {
                    // Determine scaling to fit within 500x500 while preserving aspect ratio
                    const int targetSize = 500;
                    double scale = Math.Min((double)targetSize / originalBitmap.Width, (double)targetSize / originalBitmap.Height);
                    int newWidth = (int)(originalBitmap.Width * scale);
                    int newHeight = (int)(originalBitmap.Height * scale);

                    // Create a new bitmap with padding (white background)
                    using (Bitmap paddedBitmap = new Bitmap(targetSize, targetSize))
                    using (Graphics graphics = Graphics.FromImage(paddedBitmap))
                    {
                        graphics.Clear(Color.White);
                        int offsetX = (targetSize - newWidth) / 2;
                        int offsetY = (targetSize - newHeight) / 2;
                        graphics.DrawImage(originalBitmap, new Rectangle(offsetX, offsetY, newWidth, newHeight));

                        // Save the resized image
                        string outputImagePath = $"extracted_resized_{index}.png";
                        paddedBitmap.Save(outputImagePath);
                        extractedCount++;
                    }
                }
            }

            index++;
        }

        // Validate that at least one image was processed
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted and resized.");

        // Clean up temporary files (optional)
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
    }
}
