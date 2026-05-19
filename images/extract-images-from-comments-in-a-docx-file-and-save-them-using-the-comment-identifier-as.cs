using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ExtractCommentImages
{
    public static void Main()
    {
        // Prepare working folders
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string imagesDir = Path.Combine(workDir, "Images");
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(Path.Combine(workDir, "Output"));

        // 1. Create a deterministic sample image (sample.png)
        string sampleImagePath = Path.Combine(imagesDir, "sample.png");
        CreateSampleImage(sampleImagePath, 100, 100);

        // 2. Build a DOCX with a comment that contains the image
        string docPath = Path.Combine(workDir, "CommentImage.docx");
        CreateDocumentWithCommentImage(docPath, sampleImagePath);

        // 3. Load the document and extract images from comments
        Document doc = new Document(docPath);
        ExtractImagesFromComments(doc, Path.Combine(workDir, "Output"));
    }

    // Creates a simple white PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            // Optionally draw deterministic content here
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }

    // Creates a document, adds a paragraph and a comment that contains the image
    private static void CreateDocumentWithCommentImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some regular text
        builder.Writeln("This paragraph will have a comment with an image.");

        // Create a comment node
        Comment comment = new Comment(doc, "Author", "A", DateTime.Now);

        // Create a paragraph inside the comment
        Paragraph commentParagraph = new Paragraph(doc);
        comment.AppendChild(commentParagraph);

        // Create a shape that holds the image
        Shape shape = new Shape(doc, ShapeType.Image);
        shape.ImageData.SetImage(imagePath);
        shape.Width = 100;
        shape.Height = 100;

        // Append the shape to the comment's paragraph
        commentParagraph.AppendChild(shape);

        // Attach the comment to the current paragraph
        Paragraph currentParagraph = builder.CurrentParagraph;
        currentParagraph.AppendChild(comment);

        // Save the document
        doc.Save(docPath);
    }

    // Extracts images from all comments and saves them using the comment Id as part of the filename
    private static void ExtractImagesFromComments(Document doc, string outputFolder)
    {
        int extractedCount = 0;

        // Get all comment nodes in the document
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);
        foreach (Comment comment in commentNodes)
        {
            // Find all shape nodes inside the comment subtree
            NodeCollection shapeNodes = comment.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    // Determine file extension based on the image type stored in the shape
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string fileName = $"Comment-{comment.Id}{extension}";
                    string fullPath = Path.Combine(outputFolder, fileName);

                    // Save the image
                    shape.ImageData.Save(fullPath);
                    extractedCount++;
                }
            }
        }

        // Validation: ensure at least one image was extracted
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from comments.");
    }
}
