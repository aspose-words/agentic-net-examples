using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a few paragraphs and comments.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First paragraph with a comment.
        builder.Writeln("This is the first paragraph of the document.");
        Paragraph firstPara = sourceDoc.FirstSection.Body.LastParagraph;
        AddCommentToParagraph(sourceDoc, firstPara, "Alice", "A", "First comment on the first paragraph.");

        // Second paragraph with a comment.
        builder.Writeln("This is the second paragraph, containing important information.");
        Paragraph secondPara = sourceDoc.FirstSection.Body.LastParagraph;
        AddCommentToParagraph(sourceDoc, secondPara, "Bob", "B", "Second comment on the second paragraph.");

        // Ensure the layout is up‑to‑date so page numbers are correct.
        sourceDoc.UpdatePageLayout();

        // Collect layout information (page numbers) for nodes.
        LayoutCollector layoutCollector = new LayoutCollector(sourceDoc);

        // Create a new document that will hold the printable report.
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);
        reportBuilder.Writeln("Comments Report");
        reportBuilder.Writeln(new string('=', 30));
        reportBuilder.Writeln();

        // Enumerate all comments in the source document.
        var comments = sourceDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        foreach (Comment comment in comments)
        {
            // Get the page number where the comment starts.
            int pageNumber = layoutCollector.GetStartPageIndex(comment);

            // Retrieve the paragraph that contains the comment (if any).
            string paragraphText = comment.ParentParagraph?.GetText().Trim() ?? "(No paragraph)";

            // Get the comment text.
            string commentText = comment.GetText().Trim();

            // Write the information to the report.
            reportBuilder.Writeln($"Page {pageNumber}: \"{paragraphText}\"");
            reportBuilder.Writeln($"    Author : {comment.Author}");
            reportBuilder.Writeln($"    Comment: {commentText}");
            reportBuilder.Writeln();
        }

        // Save the source document and the report.
        sourceDoc.Save("SourceDocument.docx");
        reportDoc.Save("CommentsReport.docx");
    }

    // Helper method to add a comment anchored to a paragraph.
    private static void AddCommentToParagraph(Document doc, Paragraph paragraph, string author, string initial, string commentText)
    {
        // Create a new comment.
        Comment comment = new Comment(doc, author, initial, DateTime.Now);
        comment.SetText(commentText);

        // Create a comment range that surrounds a dummy run of text.
        CommentRangeStart rangeStart = new CommentRangeStart(doc, comment.Id);
        CommentRangeEnd rangeEnd = new CommentRangeEnd(doc, comment.Id);
        Run dummyRun = new Run(doc, "Commented text.");

        // Insert the range and the comment into the paragraph.
        paragraph.AppendChild(rangeStart);
        paragraph.AppendChild(dummyRun);
        paragraph.AppendChild(rangeEnd);
        paragraph.AppendChild(comment);
    }
}
