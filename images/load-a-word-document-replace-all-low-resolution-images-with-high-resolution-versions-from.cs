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
        // Define file and folder names.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        string lowResImagePath = Path.Combine(artifactsDir, "low_res.png");
        string highResImagePath = Path.Combine(artifactsDir, "high_res.png");
        string inputDocPath = Path.Combine(artifactsDir, "input.docx");
        string outputDocPath = Path.Combine(artifactsDir, "output.docx");

        // -------------------------------------------------
        // 1. Create sample low‑resolution and high‑resolution images.
        // -------------------------------------------------
        CreateSampleImage(lowResImagePath, 100, 100, Aspose.Drawing.Color.LightGray);   // 100×100 pixels
        CreateSampleImage(highResImagePath, 500, 500, Aspose.Drawing.Color.LightBlue); // 500×500 pixels

        // -------------------------------------------------
        // 2. Build a Word document that contains low‑resolution images.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the low‑resolution image three times.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertParagraph();
            builder.InsertImage(lowResImagePath);
        }

        doc.Save(inputDocPath);

        // -------------------------------------------------
        // 3. Load the document and replace low‑resolution images.
        // -------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);

        // Threshold: images with width less than 200 pixels are considered low‑resolution.
        const int widthThreshold = 200;
        bool anyReplaced = false;

        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                              .OfType<Shape>()
                              .Where(s => s.HasImage);

        foreach (Shape shape in shapes)
        {
            // Determine the pixel width of the current image.
            int imageWidth = shape.ImageData.ImageSize.WidthPixels;

            if (imageWidth < widthThreshold)
            {
                // Replace the image with the high‑resolution version.
                shape.ImageData.SetImage(highResImagePath);
                anyReplaced = true;
            }
        }

        if (!anyReplaced)
            throw new InvalidOperationException("No low‑resolution images were found to replace.");

        // -------------------------------------------------
        // 4. Save the modified document.
        // -------------------------------------------------
        loadedDoc.Save(outputDocPath);

        // Simple validation that the output file exists.
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("The output document was not created.", outputDocPath);
    }

    // Helper method to create a deterministic bitmap and save it to a file.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color background)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(background);
            }

            // Ensure the bitmap is saved before disposing.
            bitmap.Save(filePath);
        }
    }
}
