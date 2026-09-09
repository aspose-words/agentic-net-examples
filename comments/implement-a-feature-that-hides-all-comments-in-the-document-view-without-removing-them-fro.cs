using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some paragraphs to the document.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");

        // Create a comment for the first paragraph.
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("Review this paragraph.");
        // Attach the comment to the first paragraph.
        doc.FirstSection.Body.Paragraphs[0].AppendChild(comment1);

        // Create a second comment for the second paragraph.
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("Consider rephrasing.");
        // Attach the second comment.
        doc.FirstSection.Body.Paragraphs[1].AppendChild(comment2);

        // Hide all comments in the document view without removing them.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.Hide;
        // Rebuild the layout after changing the display mode.
        doc.UpdatePageLayout();

        // Save the document; comments remain in the file but are hidden in the view.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HiddenComments.docx");
        doc.Save(outputPath);
    }
}
