using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

namespace CommentImageExtraction
{
    public class Program
    {
        public static void Main()
        {
            // Create a deterministic sample image file.
            const string sampleImagePath = "sample.png";
            CreateSampleImage(sampleImagePath);

            // Create a new document and add a paragraph.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This paragraph contains a comment with an image.");

            // Create a comment.
            Comment comment = new Comment(doc, "Author", "A", DateTime.Now);
            Paragraph commentParagraph = new Paragraph(doc);
            comment.AppendChild(commentParagraph);

            // Create an image shape inside the comment.
            Shape imageShape = new Shape(doc, ShapeType.Image);
            imageShape.ImageData.SetImage(sampleImagePath);
            imageShape.Width = 100;
            imageShape.Height = 100;
            commentParagraph.AppendChild(imageShape);

            // Append the comment to the document (attach to the last paragraph).
            Paragraph lastParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, doc.GetChildNodes(NodeType.Paragraph, true).Count - 1, true);
            lastParagraph.AppendChild(comment);

            // Save the document.
            const string docPath = "CommentImageDoc.docx";
            doc.Save(docPath);

            // Extract images from comments.
            List<string> extractedFiles = new List<string>();
            NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);
            foreach (Comment cmnt in commentNodes)
            {
                NodeCollection shapeNodes = cmnt.GetChildNodes(NodeType.Shape, true);
                foreach (Shape shp in shapeNodes)
                {
                    if (shp.HasImage)
                    {
                        string outputFileName = $"comment-{cmnt.Id}.png";
                        shp.ImageData.Save(outputFileName);
                        extractedFiles.Add(outputFileName);
                    }
                }
            }

            // Validate that at least one image was extracted.
            if (extractedFiles.Count == 0)
                throw new InvalidOperationException("No images were extracted from comments.");

            // Optional: indicate success.
            Console.WriteLine($"Extracted {extractedFiles.Count} image(s) from comments.");
        }

        private static void CreateSampleImage(string path)
        {
            const int width = 100;
            const int height = 100;
            using (Bitmap bitmap = new Bitmap(width, height))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Aspose.Drawing.Color.White);
                    using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Red, 3))
                    {
                        graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                    }
                }
                bitmap.Save(path);
            }
        }
    }
}
