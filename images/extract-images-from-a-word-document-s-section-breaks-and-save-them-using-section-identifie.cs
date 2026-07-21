using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output folders.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create deterministic sample images using Aspose.Drawing.
        // -----------------------------------------------------------------
        string imgPath1 = Path.Combine(artifactsDir, "sample1.png");
        string imgPath2 = Path.Combine(artifactsDir, "sample2.png");
        CreateSamplePng(imgPath1, Aspose.Drawing.Color.LightBlue);
        CreateSamplePng(imgPath2, Aspose.Drawing.Color.LightCoral);

        // -----------------------------------------------------------------
        // 2. Build a Word document with several sections, each containing an image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Section 0
        builder.Writeln("Section 0");
        builder.InsertImage(imgPath1);
        // Insert a section break (new page) to start the next section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 1
        builder.Writeln("Section 1");
        builder.InsertImage(imgPath2);
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2 (no image, to demonstrate validation)
        builder.Writeln("Section 2 (no image)");
        // No image inserted here.

        // Save the document.
        string docPath = Path.Combine(artifactsDir, "SampleDocument.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract images per section.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        int sectionIndex = 0;
        foreach (Section section in loadedDoc.Sections)
        {
            // Collect all Shape nodes that have images inside the current section.
            NodeCollection shapes = section.Body.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"Section-{sectionIndex}-Image-{imageIndex}{extension}";
                    string imageFullPath = Path.Combine(artifactsDir, imageFileName);
                    shape.ImageData.Save(imageFullPath);
                    imageIndex++;
                }
            }

            // Validation: ensure at least one image was extracted for sections that contain images.
            if (imageIndex == 0 && SectionContainsImage(section))
            {
                throw new InvalidOperationException($"No images were extracted from section {sectionIndex}.");
            }

            sectionIndex++;
        }

        // -----------------------------------------------------------------
        // 4. Clean up temporary image files (optional).
        // -----------------------------------------------------------------
        // File.Delete(imgPath1);
        // File.Delete(imgPath2);
    }

    // Helper to create a simple PNG file with a solid background color.
    private static void CreateSamplePng(string filePath, Aspose.Drawing.Color backgroundColor)
    {
        const int width = 200;
        const int height = 200;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backgroundColor);
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Helper to determine if a section contains any Shape with an image.
    private static bool SectionContainsImage(Section section)
    {
        NodeCollection shapes = section.Body.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage)
                return true;
        }
        return false;
    }
}
