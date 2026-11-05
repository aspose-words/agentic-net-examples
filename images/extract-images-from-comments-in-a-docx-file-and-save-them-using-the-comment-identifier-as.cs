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
            const string sampleImagePath = "sampleImage.png";
            CreateSampleImage(sampleImagePath);

            // Create a new document and add a paragraph with a comment that contains the image.
            const string docPath = "sample.docx";
            CreateDocumentWithCommentImage(docPath, sampleImagePath);

            // Load the document and extract images from comments.
            ExtractImagesFromComments(docPath);
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
                    // Draw a simple rectangle for visual distinction.
                    using (var pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 3))
                    {
                        graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                    }
                }

                bitmap.Save(path);
            }

            if (!File.Exists(path))
                throw new InvalidOperationException($"Failed to create sample image at '{path}'.");
        }

        private static void CreateDocumentWithCommentImage(string docPath, string imagePath)
        {
            // Initialize a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a paragraph that will host the comment.
            builder.Writeln("This paragraph has a comment containing an image.");

            // Create a comment node.
            Comment comment = new Comment(doc, "Author", "A", DateTime.Now);

            // Create a paragraph inside the comment.
            Paragraph commentParagraph = new Paragraph(doc);
            comment.AppendChild(commentParagraph);

            // Create an image shape and set its image.
            Shape imageShape = new Shape(doc, ShapeType.Image);
            imageShape.ImageData.SetImage(imagePath);
            imageShape.Width = 100;
            imageShape.Height = 100;

            // Append the shape to the comment's paragraph.
            commentParagraph.AppendChild(imageShape);

            // Append the comment to the last paragraph of the document.
            Paragraph lastParagraph = doc.LastSection.Body.LastParagraph;
            lastParagraph.AppendChild(comment);

            // Save the document.
            doc.Save(docPath);

            if (!File.Exists(docPath))
                throw new InvalidOperationException($"Failed to save document at '{docPath}'.");
        }

        private static void ExtractImagesFromComments(string docPath)
        {
            // Load the document.
            Document doc = new Document(docPath);

            // Collect extracted image file names for validation.
            List<string> extractedFiles = new List<string>();

            // Enumerate all comment nodes.
            NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);
            foreach (Comment comment in commentNodes)
            {
                // Find all shape nodes within the comment.
                NodeCollection shapeNodes = comment.GetChildNodes(NodeType.Shape, true);
                foreach (Shape shape in shapeNodes)
                {
                    if (shape.HasImage)
                    {
                        // Use the comment's Id as part of the filename.
                        string outputFileName = $"comment-{comment.Id}.png";
                        shape.ImageData.Save(outputFileName);
                        extractedFiles.Add(outputFileName);
                    }
                }
            }

            // Validate that at least one image was extracted.
            if (extractedFiles.Count == 0)
                throw new InvalidOperationException("No images were extracted from comments.");

            // Optional: verify that files exist.
            foreach (string file in extractedFiles)
            {
                if (!File.Exists(file))
                    throw new InvalidOperationException($"Extracted image file not found: {file}");
            }
        }
    }
}
