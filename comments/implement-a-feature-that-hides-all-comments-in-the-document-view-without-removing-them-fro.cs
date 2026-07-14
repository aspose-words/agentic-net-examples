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
        builder.Writeln("This paragraph will have a comment attached to it.");

        // Create a comment and attach it to the current paragraph.
        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
        comment.SetText("This is a sample comment that will remain in the file.");
        builder.CurrentParagraph.AppendChild(comment);

        // Save the document with visible comments (DOCX format).
        doc.Save("OriginalWithComments.docx");

        // Hide comments in the rendered view.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.Hide;

        // Rebuild the layout after changing layout options.
        doc.UpdatePageLayout();

        // Save the document as PDF – comments will not be rendered.
        doc.Save("HiddenComments.pdf");

        // Save the document again as DOCX – comments are still present in the file.
        doc.Save("HiddenComments.docx");
    }
}
