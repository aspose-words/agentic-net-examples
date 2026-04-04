using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;
using System.Linq;

namespace CommentWithHyperlinkExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a paragraph that will hold the comment.
            builder.Writeln("This paragraph will have a comment containing a hyperlink.");

            // Create a top‑level comment.
            Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
            // Append the comment to the current paragraph.
            builder.CurrentParagraph.AppendChild(comment);

            // Move the builder inside the comment so we can add content to it.
            builder.MoveTo(comment.AppendChild(new Paragraph(doc)));

            // Insert a hyperlink into the comment.
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;
            builder.InsertHyperlink("Aspose.Words", "https://www.aspose.com/words", false);
            builder.Font.ClearFormatting();

            // Ensure that comments are rendered as PDF annotations.
            doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;
            doc.UpdatePageLayout();

            // Save the document to PDF.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            string pdfPath = "CommentWithHyperlink.pdf";
            doc.Save(pdfPath, pdfOptions);

            // Reload the PDF and verify that the comment still contains the hyperlink text.
            Document loadedPdf = new Document(pdfPath);
            var comments = loadedPdf.GetChildNodes(NodeType.Comment, true)
                                    .OfType<Comment>()
                                    .ToList();

            foreach (Comment c in comments)
            {
                string commentText = c.GetText().Trim();
                Console.WriteLine($"Comment by {c.Author}: {commentText}");
                // Simple verification that the expected hyperlink text is present.
                if (commentText.Contains("Aspose.Words"))
                {
                    Console.WriteLine("Hyperlink text verified inside the comment.");
                }
                else
                {
                    Console.WriteLine("Hyperlink text NOT found in the comment.");
                }
            }
        }
    }
}
