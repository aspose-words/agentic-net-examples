using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        const string artifactsDir = "Artifacts";
        const string outputDir = "ExtractedImages";
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image (input.png).
        // -----------------------------------------------------------------
        const string sampleImagePath = "input.png";
        // Create a 200x200 bitmap using Aspose.Drawing.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        try
        {
            // Draw a simple rectangle.
            Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
            try
            {
                g.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black))
                {
                    g.DrawRectangle(pen, 20, 20, 160, 160);
                }
            }
            finally
            {
                g.Dispose();
            }

            // Save the bitmap as PNG.
            bitmap.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }
        finally
        {
            bitmap.Dispose();
        }

        // -----------------------------------------------------------------
        // 2. Build a DOCM file that contains a few images.
        // -----------------------------------------------------------------
        string docmPath = Path.Combine(artifactsDir, "sample.docm");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the same image three times, giving each shape a distinct name.
        for (int i = 1; i <= 3; i++)
        {
            Shape shape = builder.InsertImage(sampleImagePath);
            shape.Name = $"MyImageShape_{i}";
            shape.Width = 100;
            shape.Height = 100;
        }

        // Save as DOCM.
        doc.Save(docmPath, SaveFormat.Docm);

        // -----------------------------------------------------------------
        // 3. Load the DOCM file and extract embedded images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docmPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Determine a file name based on the shape's original name.
            string baseName = !string.IsNullOrEmpty(shape.Name) ? shape.Name : $"Image_{extractedCount}";
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string outputPath = Path.Combine(outputDir, baseName + extension);

            // Save the image.
            shape.ImageData.Save(outputPath);
            extractedCount++;
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0 || Directory.GetFiles(outputDir).Length == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -----------------------------------------------------------------
        // 4. (Optional) Clean up temporary files.
        // -----------------------------------------------------------------
        // File.Delete(sampleImagePath);
        // File.Delete(docmPath);
    }
}
