using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that will contain the comment.
        builder.Writeln("This paragraph will have a custom comment.");

        // Define custom author metadata.
        const string customAuthor = "Alice Example";
        const string customInitials = "AE";

        // Create a comment with the custom author, initials, and current date/time.
        Comment comment = new Comment(doc, customAuthor, customInitials, DateTime.Now);
        comment.SetText("Please review this paragraph.");

        // Attach the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Save the document to the working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Comments.docx");
        doc.Save(outputPath);

        // Enumerate all comments and write their metadata to the console.
        var comments = doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>();
        foreach (Comment c in comments)
        {
            Console.WriteLine($"Author: {c.Author}, Initials: {c.Initial}, Date: {c.DateTime}");
            Console.WriteLine($"Text: {c.GetText().Trim()}");
        }
    }
}
