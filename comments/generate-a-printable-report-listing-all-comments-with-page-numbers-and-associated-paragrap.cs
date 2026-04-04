using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several paragraphs and comments.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Paragraph 1 with a comment.
        builder.Writeln("First paragraph of the document.");
        AddComment(builder.CurrentParagraph, "Alice", "A", DateTime.Now, "Review the opening sentence.");

        // Paragraph 2 with a comment.
        builder.Writeln("Second paragraph contains important data.");
        AddComment(builder.CurrentParagraph, "Bob", "B", DateTime.Now, "Check the figures.");

        // Paragraph 3 without a comment.
        builder.Writeln("Third paragraph is plain text.");

        // Paragraph 4 with a comment.
        builder.Writeln("Fourth paragraph needs clarification.");
        AddComment(builder.CurrentParagraph, "Charlie", "C", DateTime.Now, "Explain this point further.");

        // Ensure the layout is up‑to‑date so that page numbers are accurate.
        sourceDoc.UpdatePageLayout();

        // Save the source document (optional, just for reference).
        sourceDoc.Save("source.docx");

        // Prepare a layout collector to retrieve page numbers of nodes.
        LayoutCollector layoutCollector = new LayoutCollector(sourceDoc);

        // Gather all top‑level comments (ignore replies).
        var comments = sourceDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .Where(c => c.Ancestor == null)
                                .ToList();

        // Create a new document that will hold the printable report.
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);

        reportBuilder.Writeln("Comments Report");
        reportBuilder.Writeln(new string('=', 30));
        reportBuilder.Writeln();

        // Iterate through each comment and write its details to the report.
        foreach (Comment comment in comments)
        {
            // Determine the page number where the comment appears.
            int pageNumber = layoutCollector.GetStartPageIndex(comment);

            // Retrieve the paragraph that the comment is attached to.
            Paragraph? parentParagraph = comment.ParentParagraph;
            string paragraphText = parentParagraph?.GetText().Trim() ?? "(No paragraph)";

            // Retrieve the comment text.
            string commentText = comment.GetText().Trim();

            // Write a formatted line to the report.
            reportBuilder.Writeln($"Page {pageNumber}:");
            reportBuilder.Writeln($"  Comment: \"{commentText}\"");
            reportBuilder.Writeln($"  Paragraph: \"{paragraphText}\"");
            reportBuilder.Writeln();
        }

        // Save the report document.
        reportDoc.Save("comments_report.docx");
    }

    // Helper method to create and attach a comment to a given paragraph.
    private static void AddComment(Paragraph paragraph, string author, string initial, DateTime dateTime, string text)
    {
        // Create the comment node.
        Comment comment = new Comment(paragraph.Document, author, initial, dateTime);
        comment.SetText(text);

        // Append the comment to the paragraph.
        paragraph.AppendChild(comment);
    }
}
