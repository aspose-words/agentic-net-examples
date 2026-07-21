using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    // Simple POCO representing a comment record from a database.
    private class CommentRecord
    {
        public string Author { get; set; } = "";
        public string Initial { get; set; } = "";
        public DateTime DateTime { get; set; }
        public string Text { get; set; } = "";
        // Zero‑based index of the paragraph to which the comment will be attached.
        public int ParagraphIndex { get; set; }
    }

    public static void Main()
    {
        // Simulated database records.
        List<CommentRecord> commentData = new List<CommentRecord>
        {
            new CommentRecord
            {
                Author = "Alice Johnson",
                Initial = "AJ",
                DateTime = DateTime.Now.AddDays(-2),
                Text = "Review the introduction for clarity.",
                ParagraphIndex = 0
            },
            new CommentRecord
            {
                Author = "Bob Smith",
                Initial = "BS",
                DateTime = DateTime.Now.AddDays(-1),
                Text = "Consider adding a table here.",
                ParagraphIndex = 1
            },
            new CommentRecord
            {
                Author = "Carol Lee",
                Initial = "CL",
                DateTime = DateTime.Now,
                Text = "Typo in the last sentence.",
                ParagraphIndex = 2
            }
        };

        // Create a simple template document with three paragraphs.
        Document template = new Document();
        var builder = new DocumentBuilder(template);
        builder.Writeln("This is the first paragraph of the template.");
        builder.Writeln("This is the second paragraph of the template.");
        builder.Writeln("This is the third paragraph of the template.");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Save the template (optional, demonstrates loading).
        string templatePath = Path.Combine(outputDir, "template.docx");
        template.Save(templatePath);

        // Load the template document.
        Document doc = new Document(templatePath);

        // Insert comments based on the simulated database records.
        foreach (CommentRecord record in commentData)
        {
            // Guard against an invalid paragraph index.
            if (record.ParagraphIndex < 0 || record.ParagraphIndex >= doc.FirstSection.Body.Paragraphs.Count)
                continue;

            // Target paragraph where the comment will be attached.
            var targetParagraph = doc.FirstSection.Body.Paragraphs[record.ParagraphIndex];

            // Create a new comment node with metadata.
            var comment = new Comment(doc, record.Author, record.Initial, record.DateTime);
            comment.SetText(record.Text); // Set the comment text (creates internal paragraph(s)).

            // Append the comment to the target paragraph.
            targetParagraph.AppendChild(comment);
        }

        // Save the resulting document with comments.
        string outputPath = Path.Combine(outputDir, "document-with-comments.docx");
        doc.Save(outputPath);

        // Enumerate and display the inserted comments.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment c in comments)
        {
            Console.WriteLine($"Author: {c.Author}");
            Console.WriteLine($"Date: {c.DateTime}");
            Console.WriteLine($"Text: {c.GetText().Trim()}");
            Console.WriteLine(new string('-', 40));
        }
    }
}
