using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string imagePath = "waveform.png";
        const string docPath = "sample.docx";
        const string outputFolder = "ExtractedImages";

        // Ensure output folder exists.
        Directory.CreateDirectory(outputFolder);

        // -------------------------------------------------
        // Step 1: Create a sample waveform image using Aspose.Drawing.
        // -------------------------------------------------
        const int width = 400;
        const int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill background.
                g.Clear(Color.White);

                // Draw a simple waveform (sine-like line).
                for (int x = 0; x < width; x++)
                {
                    double radians = (double)x / width * 4 * Math.PI;
                    int y = (int)(height / 2 + Math.Sin(radians) * (height / 3));
                    bitmap.SetPixel(x, y, Color.Black);
                }
            }

            // Save the generated image to a file.
            bitmap.Save(imagePath);
        }

        // -------------------------------------------------
        // Step 2: Create a DOCX document and insert the image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // Step 3: Load the document and extract all images.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine appropriate file extension for the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outputPath = Path.Combine(outputFolder, $"extracted_{extractedCount}{extension}");

                // Save the image to the file system.
                shape.ImageData.Save(outputPath);
                extractedCount++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Optional: clean up temporary files (commented out to keep results).
        // File.Delete(imagePath);
        // File.Delete(docPath);
    }
}
