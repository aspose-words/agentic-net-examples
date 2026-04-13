using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a deterministic sample image.
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        CreateSampleImage(sampleImagePath);

        // Build a document with several sections, each containing images.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Section 0 – one image.
        builder.Writeln("Section 0");
        builder.InsertImage(sampleImagePath);
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 1 – two images.
        builder.Writeln("Section 1");
        builder.InsertImage(sampleImagePath);
        builder.InsertImage(sampleImagePath);
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2 – no images.
        builder.Writeln("Section 2");

        // Save the sample document (optional, just for reference).
        string docPath = Path.Combine(outputDir, "Sample.docx");
        doc.Save(docPath);

        // Extract images, grouping them by the section they belong to.
        int totalExtracted = 0;
        for (int secIdx = 0; secIdx < doc.Sections.Count; secIdx++)
        {
            Section section = doc.Sections[secIdx];
            NodeCollection shapes = section.Body.GetChildNodes(NodeType.Shape, true);
            int imgIdx = 0;

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string ext = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string fileName = $"Section_{secIdx}_Image_{imgIdx}{ext}";
                    string fullPath = Path.Combine(outputDir, fileName);
                    shape.ImageData.Save(fullPath);
                    imgIdx++;
                    totalExtracted++;
                }
            }
        }

        // Validation – ensure at least one image was saved.
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }

    // Generates a simple 100x100 PNG image with a black rectangle.
    private static void CreateSampleImage(string path)
    {
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black))
        {
            graphics.DrawRectangle(pen, 10, 10, 80, 80);
        }
        bitmap.Save(path);
        graphics.Dispose();
        bitmap.Dispose();
    }
}
