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
            builder.Writeln($"This is paragraph {i} containing some sample text.");

            // Create a comment anchored to the current paragraph.
            Comment comment = new Comment(sourceDoc)
            {
                Author = $"Author{i}",
                Initial = $"A{i}",
                DateTime = DateTime.Now.AddMinutes(-i * 5)
            };
            // Add visible text to the comment body.
            comment.AppendChild(new Paragraph(sourceDoc));
            comment.FirstParagraph?.AppendChild(new Run(sourceDoc, $"Comment for paragraph {i}."));

            // Append the comment to the paragraph that was just written.
            Paragraph? currentParagraph = builder.CurrentParagraph;
            currentParagraph?.AppendChild(comment);
        }

        // Save the source document (optional, for inspection).
        sourceDoc.Save("sample.docx");

        // Ensure layout is up‑to‑date so we can retrieve page numbers.
        sourceDoc.UpdatePageLayout();

        // Collector to map nodes to page numbers.
        LayoutCollector collector = new LayoutCollector(sourceDoc);

        // Gather all comments safely.
        var comments = sourceDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        // Create a new document that will hold the printable report.
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);

        reportBuilder.Writeln("Comments Report");
        reportBuilder.Writeln(new string('-', 30));
        reportBuilder.Writeln();

        foreach (Comment c in comments)
        {
            // The comment is usually a child of the paragraph it annotates.
            Paragraph? anchorParagraph = c.ParentNode as Paragraph;

            // Retrieve paragraph text safely.
            string paragraphText = anchorParagraph?.GetText().Trim() ?? "(No paragraph)";

            // Determine the page number where the paragraph starts.
            int pageNumber = anchorParagraph != null ? collector.GetStartPageIndex(anchorParagraph) : 0;

            // Write comment details to the report.
            reportBuilder.Writeln($"Page: {pageNumber}");
            reportBuilder.Writeln($"Author: {c.Author}");
            reportBuilder.Writeln($"Date: {c.DateTime:O}");
            reportBuilder.Writeln($"Paragraph Text: {paragraphText}");
            reportBuilder.Writeln($"Comment Text: {c.GetText().Trim()}");
            reportBuilder.Writeln();
        }

        // Save the report document.
        reportDoc.Save("comments-report.docx");
    }
}
