using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Create a sample document with paragraphs and comments.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Add three paragraphs, each with a comment.
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is paragraph {i}.");
            Paragraph paragraph = builder.CurrentParagraph;

            // Create a top‑level comment.
            Comment comment = new Comment(sourceDoc, $"Author{i}", $"A{i}", DateTime.Now);
            comment.SetText($"Comment for paragraph {i}.");
            paragraph.AppendChild(comment);

            // Add a reply to the comment for demonstration.
            comment.AddReply($"ReplyAuthor{i}", $"R{i}", DateTime.Now, $"Reply to comment {i}.");
        }

        // Ensure the layout is up to date so we can retrieve page numbers.
        sourceDoc.UpdatePageLayout();

        // Collector maps nodes to their page numbers.
        LayoutCollector collector = new LayoutCollector(sourceDoc);

        // Prepare the report document.
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);
        reportBuilder.Writeln("Comments Report");
        reportBuilder.Writeln("----------------");
        reportBuilder.Writeln();

        // Enumerate all comments (including replies).
        var allComments = sourceDoc.GetChildNodes(NodeType.Comment, true)
                                   .OfType<Comment>()
                                   .ToList();

        foreach (Comment comment in allComments)
        {
            // Determine the top‑level comment that anchors the comment to the document.
            Comment topComment = comment.Ancestor ?? comment;

            // The paragraph that contains the top‑level comment.
            Paragraph? paragraph = topComment.ParentNode as Paragraph;

            // Retrieve page number if the paragraph exists.
            int pageNumber = paragraph != null ? collector.GetStartPageIndex(paragraph) : -1;

            // Prepare display strings.
            string commentText = comment.GetText().Trim();
            string paragraphText = paragraph?.GetText().Trim() ?? "[No paragraph]";
            string pageInfo = pageNumber > 0 ? $"Page {pageNumber}" : "Page N/A";

            // Write the entry to the report.
            reportBuilder.Writeln($"{pageInfo}: \"{commentText}\"");
            reportBuilder.Writeln($"    Associated paragraph: \"{paragraphText}\"");
            reportBuilder.Writeln();
        }

        // Save the source document and the report.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        sourceDoc.Save(Path.Combine(outputDir, "SourceDocument.docx"));
        reportDoc.Save(Path.Combine(outputDir, "CommentsReport.docx"));
    }
}
