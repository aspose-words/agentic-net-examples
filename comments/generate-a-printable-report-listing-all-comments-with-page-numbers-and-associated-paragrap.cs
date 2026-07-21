using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Create a sample source document with paragraphs and comments.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Add several paragraphs. Every second paragraph gets a comment.
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"This is paragraph {i}.");

            if (i % 2 == 0)
            {
                // Create a comment with author metadata.
                Comment comment = new Comment(sourceDoc, $"Author{i}", $"A{i}", DateTime.Now);
                comment.SetText($"Comment for paragraph {i}.");

                // Append the comment to the current paragraph.
                builder.CurrentParagraph.AppendChild(comment);
            }
        }

        // Ensure the layout is up‑to‑date so that page numbers are correct.
        sourceDoc.UpdatePageLayout();

        // Map nodes to page numbers.
        LayoutCollector collector = new LayoutCollector(sourceDoc);

        // Gather all top‑level comments (comments that are not replies).
        var comments = sourceDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .Where(c => c.Ancestor == null)
                                .ToList();

        // -----------------------------------------------------------------
        // Create a new document that will hold the printable comments report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);

        // Report title.
        reportBuilder.Font.Size = 16;
        reportBuilder.Font.Bold = true;
        reportBuilder.Writeln("Comments Report");
        reportBuilder.Font.Size = 12;
        reportBuilder.Font.Bold = false;
        reportBuilder.Writeln();

        // List each comment with its details.
        for (int idx = 0; idx < comments.Count; idx++)
        {
            Comment comment = comments[idx];

            // Page number where the comment starts.
            int pageNumber = collector.GetStartPageIndex(comment);

            // Text of the comment.
            string commentText = comment.GetText().Trim();

            // The paragraph that the comment is attached to.
            Paragraph? para = comment.ParentParagraph;
            string paragraphText = para != null ? para.GetText().Trim() : "<No paragraph>";

            // Write the information to the report.
            reportBuilder.Writeln($"Comment {idx + 1}:");
            reportBuilder.Writeln($"  Author   : {comment.Author}");
            reportBuilder.Writeln($"  Date     : {comment.DateTime}");
            reportBuilder.Writeln($"  Page     : {pageNumber}");
            reportBuilder.Writeln($"  Text     : {commentText}");
            reportBuilder.Writeln($"  Paragraph: {paragraphText}");
            reportBuilder.Writeln(); // Blank line between entries.
        }

        // Save the report to the working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentsReport.docx");
        reportDoc.Save(outputPath);
        Console.WriteLine($"Report saved to: {outputPath}");
    }
}
