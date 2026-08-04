using System;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph will have a comment attached to it.");

        // Create a comment with author metadata.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("Review this paragraph for clarity.");

        // Append the comment to the current paragraph so it is anchored to the text.
        Paragraph? currentParagraph = builder.CurrentParagraph;
        if (currentParagraph != null)
        {
            currentParagraph.AppendChild(comment);
        }

        // Configure the layout to render comments as PDF annotations.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;

        // Rebuild the layout after changing the option.
        doc.UpdatePageLayout();

        // Save the document as PDF; comments will appear as visible annotations.
        doc.Save("DocumentWithComments.pdf", SaveFormat.Pdf);
    }
}
