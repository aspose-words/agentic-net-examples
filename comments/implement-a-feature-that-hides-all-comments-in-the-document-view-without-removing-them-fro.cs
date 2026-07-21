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

        // Add a paragraph of text.
        builder.Writeln("This is a paragraph that contains a comment.");

        // Create a comment and attach it to the current paragraph.
        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
        comment.SetText("This comment will remain in the file but be hidden in the view.");
        builder.CurrentParagraph.AppendChild(comment);

        // Save the document with visible comments (DOCX format).
        doc.Save("DocumentWithComments.docx");

        // Hide comments in the rendered view.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.Hide;

        // Rebuild the layout after changing the display mode.
        doc.UpdatePageLayout();

        // Save the document where comments are hidden (PDF format demonstrates the effect).
        doc.Save("DocumentCommentsHidden.pdf");

        // Save the DOCX again; the comments are still present in the file,
        // but the layout option will affect how they are rendered in viewers that respect it.
        doc.Save("DocumentWithComments_HiddenInView.docx");
    }
}
