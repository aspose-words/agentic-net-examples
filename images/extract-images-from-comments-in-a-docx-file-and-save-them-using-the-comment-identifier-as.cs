using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Deterministic file names.
        const string imagePath = "input.png";
        const string docPath = "CommentsWithImages.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample image file (input.png) using Aspose.Drawing.
        // -----------------------------------------------------------------
        const int imgWidth = 100;
        const int imgHeight = 100;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            // Additional drawing can be added here if desired.
            bitmap.Save(imagePath);
        }

        // -----------------------------------------------------------------
        // 2. Build a DOCX document that contains a comment with the image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph that will hold the comment.
        builder.Writeln("Paragraph with a comment that contains an image.");

        // Create a comment node (no DocumentBuilder.StartComment/EndComment).
        Comment comment = new Comment(doc, "Author", "A", DateTime.Now);

        // Every comment must contain at least one paragraph.
        Paragraph commentParagraph = new Paragraph(doc);
        comment.AppendChild(commentParagraph);

        // Create a shape that holds the image.
        Shape imageShape = new Shape(doc, ShapeType.Image);
        imageShape.ImageData.SetImage(imagePath);
        imageShape.Width = imgWidth;
        imageShape.Height = imgHeight;

        // Append the shape to the comment's paragraph.
        commentParagraph.AppendChild(imageShape);

        // Append the comment to the paragraph that was just created by the builder.
        builder.CurrentParagraph.AppendChild(comment);

        // Save the document.
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract images from comments.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection commentNodes = loadedDoc.GetChildNodes(NodeType.Comment, true);

        int extractedCount = 0;

        foreach (Comment cmnt in commentNodes)
        {
            // Enumerate all shape nodes inside the comment subtree.
            NodeCollection shapeNodes = cmnt.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    // Determine file extension based on the image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    // Use the comment identifier (Id) as part of the file name.
                    string outputFileName = $"comment-{cmnt.Id}{extension}";
                    shape.ImageData.Save(outputFileName);
                    extractedCount++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new Exception("No images were extracted from comments.");

        // Inform the user (no interactive input required).
        Console.WriteLine($"Extracted {extractedCount} image(s) from comments.");
    }
}
