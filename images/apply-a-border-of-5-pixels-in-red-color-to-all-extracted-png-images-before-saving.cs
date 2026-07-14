using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a deterministic PNG image using Aspose.Drawing
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                g.DrawRectangle(new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 3), 20, 20, 160, 160);
            }
            bitmap.Save(sampleImagePath);
        }

        // 2. Insert the PNG image into a Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        string docPath = Path.Combine(outputDir, "DocumentWithImage.docx");
        doc.Save(docPath);

        // 3. Reload the document (simulating extraction scenario)
        Document loadedDoc = new Document(docPath);

        // 4. Extract PNG images, add a 5‑pixel red border, and save them
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Get the original image bytes
            byte[] imageBytes = shape.ImageData.ToByteArray();

            using (MemoryStream ms = new MemoryStream(imageBytes))
            using (Aspose.Drawing.Bitmap bmp = new Aspose.Drawing.Bitmap(ms))
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bmp))
            using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Red, 5))
            {
                // Draw a rectangle border inside the image bounds
                g.DrawRectangle(pen, 0, 0, bmp.Width - 1, bmp.Height - 1);

                // Save the modified image
                string imageFileName = $"extracted_{imageIndex}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}";
                string imageFullPath = Path.Combine(outputDir, imageFileName);
                bmp.Save(imageFullPath);
            }

            imageIndex++;
        }

        // Validation: ensure at least one image was saved
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were extracted and saved.");
    }
}
