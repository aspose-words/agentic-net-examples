using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

namespace ExtractCommentImages
{
    public class Program
    {
        public static void Main()
        {
            // Folder for all generated files.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create a sample image that will be placed inside a comment.
            // -----------------------------------------------------------------
            string sampleImagePath = Path.Combine(outputDir, "sample.png");
            const int imgWidth = 100;
            const int imgHeight = 100;

            using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.LightBlue);
                // Simple deterministic drawing – a filled rectangle.
                graphics.FillRectangle(
                    new SolidBrush(Aspose.Drawing.Color.DarkBlue),
                    10, 10, imgWidth - 20, imgHeight - 20);
                bitmap.Save(sampleImagePath);
            }

            // -----------------------------------------------------------------
            // 2. Build a DOCX document that contains a comment with the image.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph that will host the comment.
            builder.Writeln("Paragraph with a comment that contains an image.");

            // Create the comment node.
            Comment comment = new Comment(doc, "Author", "A", DateTime.Now);

            // The comment must contain a paragraph.
            Paragraph commentParagraph = new Paragraph(doc);
            comment.AppendChild(commentParagraph);

            // Create a shape that holds the image.
            Shape imageShape = new Shape(doc, ShapeType.Image);
            imageShape.ImageData.SetImage(sampleImagePath);
            imageShape.Width = imgWidth;
            imageShape.Height = imgHeight;

            // Append the shape to the comment's paragraph.
            commentParagraph.AppendChild(imageShape);

            // Attach the comment to the previously created paragraph.
            builder.CurrentParagraph.AppendChild(comment);

            // Save the document.
            string docPath = Path.Combine(outputDir, "DocumentWithComment.docx");
            doc.Save(docPath);

            // -----------------------------------------------------------------
            // 3. Load the document and extract images from all comments.
            // -----------------------------------------------------------------
            Document loadedDoc = new Document(docPath);
            NodeCollection commentNodes = loadedDoc.GetChildNodes(NodeType.Comment, true);

            int extractedCount = 0;
            int commentIndex = 0;

            foreach (Comment cmnt in commentNodes)
            {
                // Find all shape nodes inside the comment.
                NodeCollection shapeNodes = cmnt.GetChildNodes(NodeType.Shape, true);
                int shapeIndex = 0;

                foreach (Shape shape in shapeNodes)
                {
                    if (shape.HasImage)
                    {
                        string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                        string imageFileName = $"comment-{commentIndex}_{shapeIndex}{extension}";
                        string imagePath = Path.Combine(outputDir, imageFileName);
                        shape.ImageData.Save(imagePath);
                        extractedCount++;
                        shapeIndex++;
                    }
                }

                commentIndex++;
            }

            // Validate that at least one image was extracted.
            if (extractedCount == 0)
                throw new InvalidOperationException("No images were extracted from comments.");

            // Optional: indicate completion.
            Console.WriteLine($"Extraction complete. {extractedCount} image(s) saved to \"{outputDir}\".");
        }
    }
}
