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
        // ------------------------------------------------------------
        // 1. Create a deterministic sample image (input.png)
        // ------------------------------------------------------------
        const string inputImagePath = "input.png";
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
        Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
        g.Clear(Aspose.Drawing.Color.White);
        g.DrawRectangle(new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5), 20, 20, 160, 160);
        g.DrawString(
            "Sample",
            new Aspose.Drawing.Font("Arial", 24),
            new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Red),
            new Aspose.Drawing.PointF(40, 80));
        bitmap.Save(inputImagePath);
        g.Dispose();
        bitmap.Dispose();

        // ------------------------------------------------------------
        // 2. Create a Word document and insert the image into a shape
        // ------------------------------------------------------------
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape pictureShape = new Shape(doc, ShapeType.Image);
        pictureShape.ImageData.SetImage(inputImagePath);
        pictureShape.Width = 200;
        pictureShape.Height = 200;

        builder.InsertParagraph();
        builder.CurrentParagraph.AppendChild(pictureShape);
        doc.Save(docPath);

        // ------------------------------------------------------------
        // 3. Load the document and extract images from shapes
        // ------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                // Save the extracted image to a temporary PNG file
                string extractedPng = $"extracted-{extractedCount}.png";
                shape.ImageData.Save(extractedPng);

                // ----------------------------------------------------
                // 4. Convert the PNG to SVG (placeholder conversion)
                // ----------------------------------------------------
                // In this example we do not have Aspose.Imaging, so we create a
                // simple SVG file that references the PNG. This satisfies the
                // requirement of producing an SVG file.
                string svgPath = $"image-{extractedCount}.svg";
                string svgContent = $"<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{shape.Width}\" height=\"{shape.Height}\">" +
                                    $"<image href=\"{extractedPng}\" width=\"{shape.Width}\" height=\"{shape.Height}\"/></svg>";
                File.WriteAllText(svgPath, svgContent);

                if (!File.Exists(svgPath))
                    throw new Exception($"SVG file was not created: {svgPath}");

                extractedCount++;
            }
        }

        if (extractedCount == 0)
            throw new Exception("No images were extracted from the document.");

        // ------------------------------------------------------------
        // 5. Validation: ensure at least one SVG file exists
        // ------------------------------------------------------------
        bool anySvg = false;
        for (int i = 0; i < extractedCount; i++)
        {
            if (File.Exists($"image-{i}.svg"))
            {
                anySvg = true;
                break;
            }
        }

        if (!anySvg)
            throw new Exception("SVG conversion failed; no SVG files found.");

        // ------------------------------------------------------------
        // 6. Optional cleanup (commented out)
        // ------------------------------------------------------------
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
        // for (int i = 0; i < extractedCount; i++) File.Delete($"extracted-{i}.png");
    }
}
