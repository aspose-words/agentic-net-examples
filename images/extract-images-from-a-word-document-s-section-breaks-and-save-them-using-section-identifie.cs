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
        // Create a deterministic sample image file.
        const string sampleImagePath = "sample.png";
        CreateSampleImage(sampleImagePath, 100, 100, Aspose.Drawing.Color.LightBlue);

        // Build a document with three sections, each containing the sample image.
        const string docPath = "sample.docx";
        BuildDocumentWithSectionImages(docPath, sampleImagePath);

        // Extract images from each section and save them using section identifiers.
        ExtractSectionImages(docPath);
    }

    private static void CreateSampleImage(string path, int width, int height, Aspose.Drawing.Color background)
    {
        // Create bitmap and graphics objects explicitly.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(background);
        bitmap.Save(path, Aspose.Drawing.Imaging.ImageFormat.Png);
        graphics.Dispose();
        bitmap.Dispose();

        if (!File.Exists(path))
            throw new Exception($"Failed to create sample image at '{path}'.");
    }

    private static void BuildDocumentWithSectionImages(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section image.
        builder.InsertImage(imagePath);
        // Section break.
        builder.InsertBreak(Aspose.Words.BreakType.SectionBreakNewPage);

        // Second section image.
        builder.InsertImage(imagePath);
        // Section break.
        builder.InsertBreak(Aspose.Words.BreakType.SectionBreakNewPage);

        // Third section image.
        builder.InsertImage(imagePath);

        doc.Save(docPath);

        if (!File.Exists(docPath))
            throw new Exception($"Failed to save document at '{docPath}'.");
    }

    private static void ExtractSectionImages(string docPath)
    {
        Document doc = new Document(docPath);
        int sectionCount = doc.Sections.Count;

        for (int i = 0; i < sectionCount; i++)
        {
            Section section = doc.Sections[i];
            NodeCollection shapeNodes = section.Body.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    string outputFileName = $"section-{i + 1}-image-{imageIndex + 1}.png";
                    shape.ImageData.Save(outputFileName);
                    if (!File.Exists(outputFileName))
                        throw new Exception($"Failed to save extracted image '{outputFileName}'.");
                    imageIndex++;
                }
            }

            if (imageIndex == 0)
                throw new Exception($"No images found in section {i + 1}.");
        }
    }
}
