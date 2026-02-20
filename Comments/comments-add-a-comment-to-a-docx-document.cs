using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World!");

        // Create a comment with author, initials, and timestamp.
        Comment comment = new Comment(doc, "John Doe", "J.D.", DateTime.Now);
        // Set the comment text.
        comment.SetText("This is a sample comment.");

        // Attach the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Save the document to a DOCX file.
        doc.Save("CommentExample.docx");
    }
}
