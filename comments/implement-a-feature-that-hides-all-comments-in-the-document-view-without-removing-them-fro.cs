using System;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain a comment.
        builder.Writeln("This paragraph has a comment attached to it.");

        // Create a comment, set its metadata, and add some text.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("Review the wording of this paragraph.");

        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Hide comments in the rendered output while keeping them in the file.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.Hide;
        // Rebuild the layout after changing the display mode.
        doc.UpdatePageLayout();

        // Save the document in its native format (comments are preserved).
        doc.Save("DocumentWithHiddenComments.docx");
        // Save to PDF to demonstrate that comments are not rendered.
        doc.Save("DocumentWithHiddenComments.pdf");
    }
}
