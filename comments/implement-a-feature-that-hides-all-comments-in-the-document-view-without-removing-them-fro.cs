using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add first paragraph.
        builder.Writeln("This is the first paragraph.");

        // Insert a comment attached to the first paragraph.
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("Comment on the first paragraph.");
        builder.CurrentParagraph.AppendChild(comment1);

        // Add second paragraph.
        builder.Writeln("This is the second paragraph.");

        // Insert another comment.
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("Comment on the second paragraph.");
        builder.CurrentParagraph.AppendChild(comment2);

        // Save the original document (comments are preserved).
        string originalPath = Path.Combine(outputDir, "Original.docx");
        doc.Save(originalPath);

        // Hide comments in the rendered output.
        // This does not delete the comments; they remain in the document.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.Hide;
        doc.UpdatePageLayout();

        // Save the document as PDF where comments are not rendered.
        string pdfPath = Path.Combine(outputDir, "HiddenComments.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Also save the DOCX after setting the hide mode to demonstrate that the file still contains comments.
        string hiddenDocPath = Path.Combine(outputDir, "WithHiddenMode.docx");
        doc.Save(hiddenDocPath);
    }
}
