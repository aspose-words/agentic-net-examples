using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Aspose.Drawing provides Bitmap, Graphics, Color

public class Program
{
    public static void Main()
    {
        // Folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a deterministic sample image.
        string sampleImagePath = Path.Combine(artifactsDir, "input.png");
        CreateSampleImage(sampleImagePath);

        // Create a DOCX that contains a comment with the image.
        string docPath = Path.Combine(artifactsDir, "CommentImage.docx");
        CreateDocumentWithCommentImage(docPath, sampleImagePath);

        // Extract images from all comments and save them using the comment Id as filename.
        ExtractImagesFromComments(docPath, artifactsDir);
    }

    // Generates a simple white PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath)
    {
        const int width = 200;
        const int height = 200;

        // Explicit Aspose.Drawing types to avoid System.Drawing usage.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        bitmap.Save(filePath);
        // Clean up.
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Builds a document that contains a comment; the comment holds a paragraph with an image shape.
    private static void CreateDocumentWithCommentImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some visible text.
        builder.Writeln("Paragraph with a comment that contains an image.");

        // Create the comment node.
        Comment comment = new Comment(doc, "Author", "A", DateTime.Now);

        // The comment must contain a paragraph.
        Paragraph commentParagraph = new Paragraph(doc);
        comment.AppendChild(commentParagraph);

        // Create an image shape and set its image.
        Shape shape = new Shape(doc, ShapeType.Image);
        shape.ImageData.SetImage(imagePath);
        shape.Width = 100;
        shape.Height = 100;

        // Append the shape to the comment's paragraph.
        commentParagraph.AppendChild(shape);

        // Attach the comment to the current paragraph in the main document.
        builder.CurrentParagraph.AppendChild(comment);

        // Save the document.
        doc.Save(docPath);
    }

    // Finds all comment nodes, extracts any image shapes they contain, and saves the images.
    private static void ExtractImagesFromComments(string docPath, string outputDir)
    {
        Document doc = new Document(docPath);
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        int extractedImages = 0;

        foreach (Comment comment in commentNodes.OfType<Comment>())
        {
            // Use the comment's Id as part of the filename.
            int commentId = comment.Id;

            // Find all shape nodes inside the comment.
            NodeCollection shapeNodes = comment.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string outputPath = Path.Combine(outputDir, $"comment-{commentId}{extension}");
                    shape.ImageData.Save(outputPath);
                    extractedImages++;
                }
            }
        }

        if (extractedImages == 0)
            throw new Exception("No images were extracted from comments.");
    }
}
