using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Layout;

namespace AsposeWordsCommentsToXps
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a paragraph that will contain a comment.
            builder.Writeln("Paragraph with a review note.");

            // Create a comment, set its metadata and add some text to it.
            Comment comment = new Comment(doc)
            {
                Author = "Alex",
                Initial = "AL",
                DateTime = DateTime.Now
            };
            // The comment must contain at least one paragraph and run to be visible.
            comment.AppendChild(new Paragraph(doc));
            comment.FirstParagraph?.AppendChild(new Run(doc, "Please review this paragraph."));

            // Attach the comment to the first paragraph of the document.
            Paragraph? targetParagraph = doc.FirstSection?.Body?.FirstParagraph;
            if (targetParagraph != null)
                targetParagraph.AppendChild(comment);

            // Ensure that comments are rendered as markup annotations.
            doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;
            doc.UpdatePageLayout();

            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string xpsPath = Path.Combine(outputDir, "DocumentWithComments.xps");

            // Save the document to XPS format.
            XpsSaveOptions saveOptions = new XpsSaveOptions();
            doc.Save(xpsPath, saveOptions);

            // Optional: enumerate comments and write their details to the console.
            var comments = doc.GetChildNodes(NodeType.Comment, true)
                              .OfType<Comment>()
                              .ToList();

            foreach (Comment c in comments)
            {
                Console.WriteLine($"Author: {c.Author}");
                Console.WriteLine($"Date: {c.DateTime}");
                Console.WriteLine($"Text: {c.GetText().Trim()}");
                Console.WriteLine();
            }

            Console.WriteLine($"Document saved to: {xpsPath}");
        }
    }
}
