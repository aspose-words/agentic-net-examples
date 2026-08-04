using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Layout;

public class CommentsReportGenerator
{
    public static void Main()
    {
        // Create a sample document with several paragraphs and comments.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        for (int i = 1; i <= 5; i++)
        {
            // Write a paragraph.
            builder.Writeln($"This is the text of paragraph {i}.");

            // Create a comment attached to the current paragraph.
            Comment comment = new Comment(sourceDoc, $"Author{i}", $"A{i}", DateTime.Now);
            comment.SetText($"This is comment {i} on paragraph {i}.");

            // Append the comment to the paragraph.
            builder.CurrentParagraph.AppendChild(comment);
        }

        // Save the source document (optional, just for inspection).
        sourceDoc.Save("SourceDocument.docx");

        // Ensure the document layout is up‑to‑date so we can retrieve page numbers.
        sourceDoc.UpdatePageLayout();
        LayoutCollector layoutCollector = new LayoutCollector(sourceDoc);

        // Retrieve all top‑level comments.
        var comments = sourceDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .Where(c => c.Ancestor == null)
                                .ToList();

        // Create a new document that will hold the printable report.
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);

        // Header for the report.
        reportBuilder.Writeln("Comments Report");
        reportBuilder.Writeln(new string('-', 30));
        reportBuilder.Writeln();

        // List each comment with its page number and the paragraph it annotates.
        foreach (Comment c in comments)
        {
            // Page number where the comment starts.
            int pageNumber = layoutCollector.GetStartPageIndex(c);

            // Paragraph that contains the comment anchor.
            Paragraph? parentParagraph = c.ParentParagraph;
            string paragraphText = parentParagraph?.GetText().Trim() ?? "<No paragraph>";

            // Comment details.
            string commentText = c.GetText().Trim();
            string author = c.Author ?? "<Unknown>";

            reportBuilder.Writeln($"Page {pageNumber}:");
            reportBuilder.Writeln($"  Paragraph: {paragraphText}");
            reportBuilder.Writeln($"  Comment by {author}: {commentText}");
            reportBuilder.Writeln();
        }

        // Save the report document.
        reportDoc.Save("CommentsReport.docx");
    }
}
