using System;
using Aspose.Words;
using Aspose.Words.Layout;

namespace HideCommentsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some text to the document.
            builder.Writeln("This is a sample paragraph with a comment.");

            // Create a comment and attach it to the current paragraph.
            Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
            comment.SetText("This is a comment that will be hidden in the view.");
            builder.CurrentParagraph.AppendChild(comment);

            // Optionally add a reply to demonstrate that replies are also hidden.
            comment.AddReply("Jane Smith", "JS", DateTime.Now, "Reply to the comment.");

            // Hide all comments in the document view. The comments remain in the file.
            doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.Hide;

            // Rebuild the layout after changing the display mode.
            doc.UpdatePageLayout();

            // Save the document. The comments are present in the file but not rendered.
            const string outputPath = "HiddenComments.docx";
            doc.Save(outputPath);
        }
    }
}
