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
        // Prepare folders for low‑ and high‑resolution images
        string lowResFolder = "LowRes";
        string highResFolder = "HighRes";
        Directory.CreateDirectory(lowResFolder);
        Directory.CreateDirectory(highResFolder);

        // Create sample low‑resolution images
        CreateSampleImage(Path.Combine(lowResFolder, "lowres1.png"), 100, 100, Aspose.Drawing.Color.LightGray);
        CreateSampleImage(Path.Combine(lowResFolder, "lowres2.png"), 120, 80, Aspose.Drawing.Color.LightBlue);

        // Create sample high‑resolution images (library)
        CreateSampleImage(Path.Combine(highResFolder, "highres1.png"), 500, 500, Aspose.Drawing.Color.DarkGray);
        CreateSampleImage(Path.Combine(highResFolder, "highres2.png"), 600, 400, Aspose.Drawing.Color.DarkBlue);

        // Build a Word document containing low‑resolution images
        string inputDocPath = "input.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with low‑resolution images:");
        builder.InsertImage(Path.Combine(lowResFolder, "lowres1.png"));
        builder.Writeln();
        builder.InsertImage(Path.Combine(lowResFolder, "lowres2.png"));
        doc.Save(inputDocPath);

        // Load the document for processing
        Document loadedDoc = new Document(inputDocPath);

        // Replace low‑resolution images with high‑resolution versions
        int replacementCount = 0;
        int highResIndex = 1;
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            if (!shape.HasImage)
                continue;

            // Width in points (1 point = 1/72 inch)
            double widthPoints = shape.ImageData.ImageSize.WidthPoints;

            // Consider images narrower than 150 points (~2 inches) as low‑resolution
            if (widthPoints < 150)
            {
                string highResPath = Path.Combine(highResFolder, $"highres{highResIndex}.png");
                if (File.Exists(highResPath))
                {
                    shape.ImageData.SetImage(highResPath);
                    replacementCount++;
                }
                highResIndex++;
            }
        }

        // Validate that at least one image was replaced
        if (replacementCount == 0)
            throw new InvalidOperationException("No low‑resolution images were found for replacement.");

        // Save the updated document
        string outputDocPath = "output.docx";
        loadedDoc.Save(outputDocPath);

        // Validate output file existence
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("The output document was not created.", outputDocPath);
    }

    // Helper method to create a deterministic sample image
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backgroundColor);
                // Draw a simple rectangle border for visual distinction
                using (Pen pen = new Pen(Aspose.Drawing.Color.Black, 2))
                {
                    graphics.DrawRectangle(pen, 0, 0, width - 1, height - 1);
                }
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
