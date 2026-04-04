using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    // Simple POCO representing a comment record that might come from a database.
    private class CommentData
    {
        public string Author { get; set; } = "";
        public string Initial { get; set; } = "";
        public string Text { get; set; } = "";
        public DateTime DateTime { get; set; }
    }

    public static void Main()
    {
        // Simulate reading comment data from a database.
        List<CommentData> commentRecords = GetCommentDataFromDatabase();

        // Create a template document in memory.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Paragraph 1: Introduction.");
        builder.Writeln("Paragraph 2: Details.");
        builder.Writeln("Paragraph 3: Conclusion.");

        // Insert each comment into the corresponding paragraph.
        // For this example we attach the first comment to the first paragraph,
        // the second comment to the second paragraph, and so on.
        for (int i = 0; i < commentRecords.Count; i++)
        {
            // Ensure we have a paragraph to attach the comment to.
            if (i >= template.FirstSection.Body.Paragraphs.Count)
                break; // No more paragraphs; stop adding comments.

            Paragraph paragraph = template.FirstSection.Body.Paragraphs[i];
            CommentData data = commentRecords[i];

            // Create a new comment node with metadata.
            Comment comment = new Comment(template, data.Author, data.Initial, data.DateTime);
            comment.SetText(data.Text);

            // Append the comment to the paragraph. This makes the comment visible in Word.
            paragraph.AppendChild(comment);
        }

        // Optional: enumerate and print all comments to the console.
        foreach (Comment c in template.GetChildNodes(NodeType.Comment, true).OfType<Comment>())
        {
            Console.WriteLine($"Author: {c.Author}, Date: {c.DateTime}, Text: {c.GetText().Trim()}");
        }

        // Save the resulting document.
        string outputPath = "CommentsResult.docx";
        template.Save(outputPath);
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }

    // Mock method that pretends to fetch comment rows from a database.
    private static List<CommentData> GetCommentDataFromDatabase()
    {
        return new List<CommentData>
        {
            new CommentData
            {
                Author = "Alice Johnson",
                Initial = "AJ",
                Text = "Please review the introduction.",
                DateTime = DateTime.Now.AddDays(-2)
            },
            new CommentData
            {
                Author = "Bob Smith",
                Initial = "BS",
                Text = "Add more details to this section.",
                DateTime = DateTime.Now.AddDays(-1)
            },
            new CommentData
            {
                Author = "Carol Lee",
                Initial = "CL",
                Text = "Conclusion looks good.",
                DateTime = DateTime.Now
            }
        };
    }
}
