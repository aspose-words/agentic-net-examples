using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with some text.
        builder.Writeln("This is a sample paragraph.");

        // Retrieve the paragraph we just added.
        Paragraph paragraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

        // Create a comment and add it to the paragraph.
        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.Runs.Add(new Run(doc, "This is a comment on the paragraph."));
        comment.AppendChild(commentParagraph);
        paragraph.AppendChild(comment);

        // Save the document.
        string outputPath = "CommentedDocument.docx";
        doc.Save(outputPath);

        // Retrieve and display the comment text.
        foreach (Comment c in doc.GetChildNodes(NodeType.Comment, true))
        {
            Console.WriteLine($"Comment by {c.Author}: {c.GetText().Trim()}");
        }
    }
}
