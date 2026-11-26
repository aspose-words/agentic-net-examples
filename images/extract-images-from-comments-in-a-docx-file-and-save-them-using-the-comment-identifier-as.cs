using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample image to be used inside a comment
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        CreateSampleImage(sampleImagePath);

        // 2. Create a DOCX document with a comment that contains the image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph will have a comment containing an image.");

        // Capture the paragraph we just wrote
        Paragraph paragraph = builder.CurrentParagraph;

        // Create the comment node
        Aspose.Words.Comment comment = new Aspose.Words.Comment(doc, "Author", "A", DateTime.Now);
        Paragraph commentParagraph = new Paragraph(doc);
        comment.AppendChild(commentParagraph);

        // Create the image shape inside the comment
        Shape imageShape = new Shape(doc, ShapeType.Image);
        imageShape.ImageData.SetImage(sampleImagePath);
        commentParagraph.AppendChild(imageShape);

        // Append the comment to the paragraph (valid container)
        paragraph.AppendChild(comment);

        // Save the document (optional, for verification)
        string docPath = Path.Combine(outputDir, "DocumentWithComment.docx");
        doc.Save(docPath);

        // 3. Extract images from comments and save them using comment identifier as filename
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);
        int extractedCount = 0;

        for (int i = 0; i < commentNodes.Count; i++)
        {
            Aspose.Words.Comment cm = (Aspose.Words.Comment)commentNodes[i];
            NodeCollection shapeNodes = cm.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = Path.Combine(outputDir, $"comment-{i + 1}.png");
                    shape.ImageData.Save(imageFileName);
                    if (!File.Exists(imageFileName))
                        throw new InvalidOperationException($"Failed to save extracted image to '{imageFileName}'.");
                    extractedCount++;
                }
            }
        }

        // Validation
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from comments.");

        // Cleanup temporary sample image (optional)
        // File.Delete(sampleImagePath);
    }

    private static void CreateSampleImage(string path)
    {
        const int width = 200;
        const int height = 100;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.LightBlue);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.DarkBlue, 3))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }
            }
            bitmap.Save(path, ImageFormat.Png);
        }
    }
}
