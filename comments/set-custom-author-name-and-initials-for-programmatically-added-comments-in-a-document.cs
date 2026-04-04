using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain a comment.
        builder.Writeln("This is a paragraph that will have a comment.");

        // Define custom author metadata for the comment.
        string author = "Alice Smith";
        string initials = "AS";
        DateTime commentDate = DateTime.Now;

        // Create the comment with the custom author, initials, and date.
        Comment comment = new Comment(doc, author, initials, commentDate);
        comment.SetText("Please review this paragraph.");

        // Attach the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Save the document to the working directory.
        string outputPath = "output.docx";
        doc.Save(outputPath);

        // Enumerate all comments in the document and output their metadata.
        var comments = doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>();
        foreach (Comment c in comments)
        {
            Console.WriteLine($"Comment Author: {c.Author}, Initials: {c.Initial}, Date: {c.DateTime}");
            Console.WriteLine($"Text: {c.GetText().Trim()}");
        }
    }
}
