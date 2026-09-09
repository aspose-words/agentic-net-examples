using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("First paragraph.");

        // Add a comment authored by Alice.
        Comment commentAlice = new Comment(doc, "Alice", "A", DateTime.Now);
        commentAlice.SetText("Please review this paragraph.");
        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(commentAlice);

        builder.Writeln("Second paragraph.");

        // Add a comment authored by Bob.
        Comment commentBob = new Comment(doc, "Bob", "B", DateTime.Now);
        commentBob.SetText("Check the spelling here.");
        builder.CurrentParagraph.AppendChild(commentBob);

        // Save the original document for reference.
        doc.Save("Original.docx");

        // -----------------------------------------------------------------
        // Process comments: keep only those authored by "Alice", remove others.
        // -----------------------------------------------------------------
        var allComments = doc.GetChildNodes(NodeType.Comment, true)
                             .OfType<Comment>()
                             .ToList();

        foreach (Comment c in allComments)
        {
            // If the comment author is not Alice, remove the comment.
            if (!string.Equals(c.Author, "Alice", StringComparison.OrdinalIgnoreCase))
            {
                c.Remove();
            }
        }

        // Save the revised document containing only accepted comments.
        doc.Save("Revised.docx");
    }
}
