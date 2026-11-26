using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // 1. Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(outputDir, "input.png");
        CreateSampleImage(sampleImagePath, 200, 100);

        // 2. Build a DOCX with two content controls, each containing the sample image
        string docPath = Path.Combine(outputDir, "sample.docx");
        CreateDocumentWithContentControls(docPath, sampleImagePath);

        // 3. Load the document and extract images from each content control,
        //    renaming the files using the control's tag identifier.
        ExtractImagesFromContentControls(docPath, outputDir);
    }

    private static void CreateSampleImage(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle for visual distinction
                g.FillRectangle(new SolidBrush(Color.LightBlue), 10, 10, width - 20, height - 20);
            }
            bitmap.Save(path, ImageFormat.Png);
        }

        if (!File.Exists(path))
            throw new Exception($"Failed to create sample image at '{path}'.");
    }

    private static void CreateDocumentWithContentControls(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First content control
        StructuredDocumentTag sdt1 = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        sdt1.Tag = "ImageControl1";
        Paragraph para1 = new Paragraph(doc);
        sdt1.AppendChild(para1);
        doc.FirstSection.Body.AppendChild(sdt1);
        builder.MoveTo(para1);
        builder.InsertImage(imagePath);

        // Second content control
        StructuredDocumentTag sdt2 = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        sdt2.Tag = "ImageControl2";
        Paragraph para2 = new Paragraph(doc);
        sdt2.AppendChild(para2);
        doc.FirstSection.Body.AppendChild(sdt2);
        builder.MoveTo(para2);
        builder.InsertImage(imagePath);

        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new Exception($"Failed to save document at '{docPath}'.");
    }

    private static void ExtractImagesFromContentControls(string docPath, string outputDir)
    {
        Document doc = new Document(docPath);
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        int totalExtracted = 0;

        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            string tag = string.IsNullOrEmpty(sdt.Tag) ? "UnnamedControl" : sdt.Tag;
            NodeCollection shapeNodes = sdt.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = $"{tag}_{imageIndex}.png";
                    string imagePath = Path.Combine(outputDir, imageFileName);
                    shape.ImageData.Save(imagePath);
                    if (!File.Exists(imagePath))
                        throw new Exception($"Failed to save extracted image to '{imagePath}'.");
                    imageIndex++;
                    totalExtracted++;
                }
            }
        }

        if (totalExtracted == 0)
            throw new Exception("No images were extracted from content controls.");
    }
}
