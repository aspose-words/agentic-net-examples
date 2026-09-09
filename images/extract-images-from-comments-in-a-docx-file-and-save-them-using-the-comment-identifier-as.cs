using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image (sample.png).
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Simple visual cue – a black rectangle.
            graphics.DrawRectangle(Pens.Black, 10, 10, 80, 80);
            bitmap.Save(sampleImagePath);
        }

        // -----------------------------------------------------------------
        // 2. Build a DOCX that contains a comment with the image inside.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Document with a comment that holds an image.");

        // Create a comment node.
        Comment comment = new Comment(doc, "Author", "A", DateTime.Now);

        // The comment must contain a paragraph.
        Paragraph commentParagraph = new Paragraph(doc);
        comment.AppendChild(commentParagraph);

        // Create an image shape and set its image.
        Shape imageShape = new Shape(doc, ShapeType.Image);
        imageShape.ImageData.SetImage(sampleImagePath);
        imageShape.Width = 100;
        imageShape.Height = 100;

        // Append the shape to the comment's paragraph.
        commentParagraph.AppendChild(imageShape);

        // Append the comment to the current paragraph in the main document.
        builder.CurrentParagraph.AppendChild(comment);

        // Save the document.
        string docPath = Path.Combine(artifactsDir, "CommentImage.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Extract images from all comments and save them using comment id.
        // -----------------------------------------------------------------
        int extractedImages = 0;
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        int commentIndex = 0;
        foreach (Comment c in commentNodes.OfType<Comment>())
        {
            commentIndex++;

            // Search for Shape nodes inside the comment subtree.
            NodeCollection shapeNodes = c.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Determine file extension based on image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    // Use comment Id if available; otherwise fallback to the sequential index.
                    string fileName = $"comment-{c.Id}{extension}";
                    string outPath = Path.Combine(artifactsDir, fileName);
                    shape.ImageData.Save(outPath);
                    extractedImages++;
                }
            }
        }

        if (extractedImages == 0)
            throw new Exception("No images were extracted from comments.");

        // Optional: indicate success.
        Console.WriteLine($"Extracted {extractedImages} image(s) from comments to folder: {artifactsDir}");
    }
}
