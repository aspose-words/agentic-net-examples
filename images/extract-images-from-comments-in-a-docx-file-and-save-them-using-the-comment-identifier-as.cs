using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample image that will be inserted into a comment.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            // Draw a simple rectangle to make the image recognizable.
            using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 2))
            {
                graphics.DrawRectangle(pen, 10, 10, 80, 80);
            }
            bitmap.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Build a DOCX document with a comment that contains the image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph with a comment that holds an image.");

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
        // 3. Load the document and extract images from all comments.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection commentNodes = loadedDoc.GetChildNodes(NodeType.Comment, true);

        int extractedImages = 0;
        foreach (Comment c in commentNodes.OfType<Comment>())
        {
            // Find all shape nodes inside the comment.
            NodeCollection shapeNodes = c.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Determine file extension based on image type.
                    string extension = Aspose.Words.FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    // Use the comment's Id as part of the filename.
                    string imageFileName = $"comment-{c.Id}{extension}";
                    string imagePath = Path.Combine(artifactsDir, imageFileName);
                    shape.ImageData.Save(imagePath);
                    extractedImages++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImages == 0)
            throw new Exception("No images were extracted from comments.");

        // Indicate completion.
        Console.WriteLine($"Extraction complete. {extractedImages} image(s) saved to '{artifactsDir}'.");
    }
}
