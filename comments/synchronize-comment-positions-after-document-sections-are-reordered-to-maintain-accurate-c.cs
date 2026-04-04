using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build two sections, each containing a paragraph with a comment.
        for (int i = 1; i <= 2; i++)
        {
            // For the second section start a new page.
            if (i > 1)
                builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Write some text that will be the anchor for the comment.
            builder.Writeln($"Section {i} main paragraph.");

            // Create a comment object.
            Comment comment = new Comment(doc, $"Author{i}", $"A{i}", DateTime.Now);
            comment.SetText($"Comment for section {i}");

            // The comment must be anchored to a range: start, text, end, then the comment node.
            Paragraph para = builder.CurrentParagraph;

            // Insert the start of the comment range.
            para.AppendChild(new CommentRangeStart(doc, comment.Id));

            // The text that is covered by the comment.
            para.AppendChild(new Run(doc, $"Commented text in section {i}."));

            // Insert the end of the comment range.
            para.AppendChild(new CommentRangeEnd(doc, comment.Id));

            // Finally, attach the comment itself.
            para.AppendChild(comment);
        }

        // Save the original document.
        const string originalPath = "original.docx";
        doc.Save(originalPath);

        // ------------------------------------------------------------
        // Reorder sections: move the second section before the first.
        // ------------------------------------------------------------
        if (doc.Sections.Count >= 2)
        {
            // Remove the second section and insert it at the beginning.
            Section second = doc.Sections[1];
            doc.Sections.RemoveAt(1);
            doc.Sections.Insert(0, second);
        }

        // ------------------------------------------------------------
        // Synchronize comment IDs after the reorder (safety check).
        // ------------------------------------------------------------
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment comment in comments)
        {
            // Find the nearest preceding CommentRangeStart.
            Node? startNode = comment.PreviousSibling;
            while (startNode != null && !(startNode is CommentRangeStart))
                startNode = startNode.PreviousSibling;

            if (startNode is CommentRangeStart rangeStart)
            {
                // Align the comment's Id with the range start.
                comment.Id = rangeStart.Id;
            }

            // Find the nearest following CommentRangeEnd.
            Node? endNode = comment.NextSibling;
            while (endNode != null && !(endNode is CommentRangeEnd))
                endNode = endNode.NextSibling;

            if (endNode is CommentRangeEnd rangeEnd && startNode is CommentRangeStart)
            {
                // Ensure the range end uses the same Id.
                rangeEnd.Id = comment.Id;
            }
        }

        // Save the reordered document.
        const string reorderedPath = "reordered.docx";
        doc.Save(reorderedPath);

        // Output a simple verification to the console.
        Console.WriteLine("Comments after reordering and synchronization:");
        foreach (Comment c in doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>())
        {
            Console.WriteLine($"Author: {c.Author}, Id: {c.Id}, Text: {c.GetText().Trim()}");
        }
    }
}
