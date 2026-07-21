using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph of text.
        builder.Writeln("This is a sample paragraph that will have a comment.");

        // Create a comment anchored to the current paragraph.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("Please review this sentence.");
        // Append the comment to the paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Ensure that comments are rendered as markup annotations.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;

        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document to XPS format.
        string xpsPath = Path.Combine(outputDir, "DocumentWithComments.xps");
        doc.Save(xpsPath, SaveFormat.Xps);

        // Optional: also save the original DOCX for reference.
        string docxPath = Path.Combine(outputDir, "DocumentWithComments.docx");
        doc.Save(docxPath, SaveFormat.Docx);
    }
}
