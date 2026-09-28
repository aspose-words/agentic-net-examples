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
        // Define paths
        string baseDir = Directory.GetCurrentDirectory();
        string sampleImagePath = Path.Combine(baseDir, "sample.png");
        string docPath = Path.Combine(baseDir, "sample.docx");
        string outputFolder = Path.Combine(baseDir, "ExtractedImages");

        // Ensure output folder exists
        Directory.CreateDirectory(outputFolder);

        // -------------------------------------------------
        // Create a deterministic sample image (sample.png)
        // -------------------------------------------------
        int imgWidth = 200;
        int imgHeight = 100;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle
                g.DrawRectangle(Pens.Black, 10, 10, imgWidth - 20, imgHeight - 20);
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // Create a DOCX document and insert the sample image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // Load the DOCX document and extract all embedded images
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                // Determine file extension based on image type
                string extension = "." + shape.ImageData.ImageType.ToString().ToLowerInvariant();
                string outputPath = Path.Combine(outputFolder, $"image-{imageIndex}{extension}");

                // Save the image to the output folder
                shape.ImageData.Save(outputPath);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (imageIndex == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        // Optional: indicate completion (no interactive prompts)
        Console.WriteLine($"Extracted {imageIndex} image(s) to \"{outputFolder}\".");
    }
}
