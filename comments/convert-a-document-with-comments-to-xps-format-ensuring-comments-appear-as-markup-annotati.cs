using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain a commented region.
        builder.Writeln("Paragraph before the commented text.");

        // Create a comment with author metadata.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("This is a sample comment that will appear as an annotation in the XPS output.");

        // Anchor the comment to a region of text.
        Paragraph para = doc.FirstSection.Body.FirstParagraph;
        para.AppendChild(new CommentRangeStart(doc, comment.Id));
        para.AppendChild(new Run(doc, "Commented text"));
        para.AppendChild(new CommentRangeEnd(doc, comment.Id));
        para.AppendChild(comment);

        // Add another paragraph after the comment.
        builder.Writeln("Paragraph after the commented text.");

        // Enumerate comments and write basic info to the console.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment c in comments)
        {
            Console.WriteLine($"{c.Author}: {c.GetText().Trim()}");
        }

        // Save the document to XPS format. Comments are rendered as markup annotations.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        doc.Save("DocumentWithComments.xps", xpsOptions);
    }
}
