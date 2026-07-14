using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper method to add a comment with a specific date.
        void AddComment(string author, string initials, DateTime date, string commentText)
        {
            // Write a paragraph that will hold the comment.
            builder.Writeln($"Paragraph for comment by {author}.");

            // Create the comment node with the supplied metadata.
            Comment comment = new Comment(doc, author, initials, date);
            comment.SetText(commentText);

            // Append the comment to the current paragraph.
            builder.CurrentParagraph.AppendChild(comment);
        }

        // Add three comments with different dates.
        AddComment("Alice", "A", DateTime.Now.AddDays(-30), "This comment is 30 days old.");
        AddComment("Bob", "B", DateTime.Now.AddDays(-10), "This comment is 10 days old.");
        AddComment("Charlie", "C", DateTime.Now, "This comment is from today.");

        // Define the age threshold: comments older than 15 days will be removed.
        DateTime threshold = DateTime.Now.AddDays(-15);

        // Enumerate all comment nodes safely and remove those older than the threshold.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList(); // Create a copy to avoid modifying the collection during iteration.

        foreach (Comment c in comments)
        {
            if (c.DateTime < threshold)
                c.Remove();
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Save the resulting document.
        string outputPath = Path.Combine(outputDir, "CommentsFiltered.docx");
        doc.Save(outputPath);
    }
}
