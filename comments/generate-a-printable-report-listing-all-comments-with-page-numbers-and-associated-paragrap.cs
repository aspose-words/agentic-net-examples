using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several paragraphs and comments.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Add three paragraphs, each with a comment.
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Paragraph {i}: This is some sample text for paragraph {i}.");

            // Create a comment anchored to the current paragraph.
            Comment comment = new Comment(sourceDoc, $"Author{i}", $"A{i}", DateTime.Now);
            comment.SetText($"Comment {i} on paragraph {i}.");

            // Append the comment to the paragraph.
            builder.CurrentParagraph.AppendChild(comment);
        }

        // Save the source document (optional, just for inspection).
        sourceDoc.Save("SourceDocument.docx");

        // Ensure layout is up‑to‑date so that page numbers are accurate.
        sourceDoc.UpdatePageLayout();

        // Use LayoutCollector to retrieve page numbers for nodes.
        LayoutCollector collector = new LayoutCollector(sourceDoc);

        // Prepare a new document that will hold the printable report.
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);

        // Write a header for the report.
        reportBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        reportBuilder.Writeln("Comments Report");
        reportBuilder.ParagraphFormat.ClearFormatting();

        // Enumerate all top‑level comments (ignore replies) and collect required data.
        var comments = sourceDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .Where(c => c.Ancestor == null) // top‑level only
                                .ToList();

        foreach (Comment comment in comments)
        {
            // The paragraph that contains the comment anchor.
            Paragraph? anchorParagraph = comment.ParentParagraph;

            // Safely obtain the paragraph text.
            string paragraphText = anchorParagraph?.GetText().Trim() ?? "(No paragraph)";

            // Retrieve the page number where the comment is located using LayoutCollector.
            int pageNumber = collector.GetStartPageIndex(comment);

            // Write the comment entry into the report.
            reportBuilder.Writeln($"Comment by {comment.Author} on {comment.DateTime:yyyy-MM-dd HH:mm}");
            reportBuilder.Writeln($"Page: {pageNumber}");
            reportBuilder.Writeln($"Comment Text: {comment.GetText().Trim()}");
            reportBuilder.Writeln($"Associated Paragraph: {paragraphText}");
            reportBuilder.Writeln(); // blank line between entries
        }

        // Save the report document.
        reportDoc.Save("CommentsReport.docx");
    }
}
