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

        // First paragraph with a comment.
        builder.Writeln("This is the first paragraph of the document.");
        AddCommentToCurrentParagraph(sourceDoc, builder, "Alice", "AL", "First comment on the first paragraph.");

        // Second paragraph with a comment.
        builder.Writeln("This is the second paragraph, containing important information.");
        AddCommentToCurrentParagraph(sourceDoc, builder, "Bob", "BO", "Second comment, reviewing the second paragraph.");

        // Third paragraph without a comment.
        builder.Writeln("This paragraph has no comments attached.");

        // Fourth paragraph with a comment.
        builder.Writeln("Final paragraph before the report generation.");
        AddCommentToCurrentParagraph(sourceDoc, builder, "Charlie", "CH", "Final comment for the report.");

        // Ensure layout is up‑to‑date so page numbers can be retrieved.
        sourceDoc.UpdatePageLayout();

        // Prepare a layout collector to obtain page numbers.
        LayoutCollector layoutCollector = new LayoutCollector(sourceDoc);

        // Enumerate all comments in the document.
        var comments = sourceDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        // Create a new document that will hold the printable report.
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);

        reportBuilder.Writeln("Comments Report");
        reportBuilder.Writeln(new string('=', 20));
        reportBuilder.Writeln();

        foreach (Comment comment in comments)
        {
            // Get the page number where the comment starts.
            int pageNumber = layoutCollector.GetStartPageIndex(comment);

            // Find the paragraph that the comment is associated with.
            Paragraph? associatedParagraph = FindAssociatedParagraph(comment);

            // Write comment details to the report.
            reportBuilder.Writeln($"Page Number : {pageNumber}");
            reportBuilder.Writeln($"Paragraph   : {(associatedParagraph != null ? associatedParagraph.GetText().Trim() : "N/A")}");
            reportBuilder.Writeln($"Author      : {comment.Author}");
            reportBuilder.Writeln($"Date        : {comment.DateTime:O}");
            reportBuilder.Writeln($"Comment Text: {comment.GetText().Trim()}");
            reportBuilder.Writeln();
        }

        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Save the source document and the generated report.
        sourceDoc.Save(Path.Combine(outputDir, "SourceDocument.docx"));
        reportDoc.Save(Path.Combine(outputDir, "CommentsReport.docx"));
    }

    // Helper method to add a comment to the paragraph that was just written.
    private static void AddCommentToCurrentParagraph(Document doc, DocumentBuilder builder, string author, string initials, string commentText)
    {
        // The paragraph that was just added is the current paragraph of the builder.
        Paragraph? paragraph = builder.CurrentParagraph;
        if (paragraph == null)
            return;

        // Create the comment node with metadata.
        Comment comment = new Comment(doc)
        {
            Author = author,
            Initial = initials,
            DateTime = DateTime.Now
        };

        // Add a paragraph and run inside the comment so it contains visible text.
        comment.AppendChild(new Paragraph(doc));
        comment.FirstParagraph?.AppendChild(new Run(doc, commentText));

        // Insert the comment as a child of the paragraph (allowed location).
        paragraph.AppendChild(comment);
    }

    // Helper method to locate the paragraph that a given comment is attached to.
    private static Paragraph? FindAssociatedParagraph(Comment comment)
    {
        // In this example the comment is a child of its associated paragraph.
        return comment.ParentNode as Paragraph;
    }
}
