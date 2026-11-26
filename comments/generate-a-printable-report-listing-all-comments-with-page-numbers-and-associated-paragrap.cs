using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Layout;

public class CommentsReportExample
{
    public static void Main()
    {
        // Create a sample document with paragraphs and comments.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First paragraph with a comment.
        builder.Writeln("This is the first paragraph of the document.");
        Comment comment1 = new Comment(sourceDoc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now.AddDays(-2)
        };
        comment1.AppendChild(new Paragraph(sourceDoc));
        comment1.FirstParagraph.AppendChild(new Run(sourceDoc, "Please review the first paragraph."));
        // Attach the comment to the paragraph that was just written.
        sourceDoc.FirstSection.Body.FirstParagraph.AppendChild(comment1);

        // Second paragraph with a comment.
        builder.Writeln("Second paragraph contains important information.");
        Comment comment2 = new Comment(sourceDoc)
        {
            Author = "Bob",
            Initial = "B",
            DateTime = DateTime.Now.AddDays(-1)
        };
        comment2.AppendChild(new Paragraph(sourceDoc));
        comment2.FirstParagraph.AppendChild(new Run(sourceDoc, "Check the data accuracy here."));
        // Attach the comment to the newly added paragraph.
        Paragraph secondPara = sourceDoc.FirstSection.Body.LastParagraph;
        secondPara.AppendChild(comment2);

        // Third paragraph without a comment.
        builder.Writeln("A paragraph without any comments.");

        // Ensure layout is up‑to‑date so page numbers are correct.
        sourceDoc.UpdatePageLayout();

        // Collector to retrieve page numbers of nodes.
        LayoutCollector collector = new LayoutCollector(sourceDoc);

        // Enumerate all comments safely.
        var comments = sourceDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        // Create a new document that will hold the printable report.
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);

        reportBuilder.Writeln("Comments Report");
        reportBuilder.Writeln("================");
        reportBuilder.Writeln();

        foreach (Comment c in comments)
        {
            // Find the paragraph that the comment is attached to.
            Paragraph? anchorParagraph = c.GetAncestor(NodeType.Paragraph) as Paragraph;
            if (anchorParagraph == null)
                continue; // Skip if we cannot locate the paragraph.

            // Retrieve the page number where the paragraph starts.
            int pageNumber = collector.GetStartPageIndex(anchorParagraph);

            // Gather required information.
            string paragraphText = anchorParagraph.GetText().Trim();
            string commentText = c.GetText().Trim();
            string author = c.Author;
            DateTime date = c.DateTime;

            // Write the information into the report.
            reportBuilder.Writeln($"Page: {pageNumber}");
            reportBuilder.Writeln($"Paragraph: {paragraphText}");
            reportBuilder.Writeln($"Author: {author}");
            reportBuilder.Writeln($"Date: {date:yyyy-MM-dd HH:mm}");
            reportBuilder.Writeln($"Comment: {commentText}");
            reportBuilder.Writeln();
        }

        // Save the source document and the report.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        sourceDoc.Save(Path.Combine(outputDir, "SourceDocument.docx"));
        reportDoc.Save(Path.Combine(outputDir, "CommentsReport.docx"));
    }
}
