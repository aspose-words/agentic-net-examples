using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph will have a comment attached to it.");

        // Create a top‑level comment with author metadata.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        // Set the comment text – this also creates a paragraph inside the comment.
        comment.SetText("Review the wording of this paragraph.");

        // Append the comment to the current paragraph so it is anchored to the text.
        builder.CurrentParagraph?.AppendChild(comment);

        // Add a reply to demonstrate comment threading (optional).
        comment.AddReply("Bob", "B", DateTime.Now, "Looks good to me.");

        // Ensure comments are rendered as balloons (markup annotations) in the XPS output.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInBalloons;
        doc.UpdatePageLayout();

        // Prepare XPS save options.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DocumentWithComments.xps");

        // Save the document to XPS format.
        doc.Save(outputPath, xpsOptions);

        // Enumerate comments and write a simple summary to the console.
        var comments = doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>();
        foreach (Comment c in comments)
        {
            Console.WriteLine($"{c.Author} ({c.DateTime}): {c.GetText().Trim()}");
        }
    }
}
