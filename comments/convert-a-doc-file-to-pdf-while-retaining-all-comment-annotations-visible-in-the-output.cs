using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph of text.
        builder.Writeln("This is a paragraph that will have a comment attached.");

        // Create a comment, set its metadata, and add text.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("Review this paragraph for clarity.");

        // Attach the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Ensure comments are rendered as PDF annotations.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;
        doc.UpdatePageLayout();

        // Save the original DOC file (optional, for verification).
        string docPath = Path.Combine(outputDir, "SampleDocument.doc");
        doc.Save(docPath, SaveFormat.Doc);

        // Convert and save the document as PDF with visible comments.
        string pdfPath = Path.Combine(outputDir, "SampleDocumentWithComments.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Output the locations of the generated files.
        Console.WriteLine($"DOC file saved to: {docPath}");
        Console.WriteLine($"PDF file saved to: {pdfPath}");
    }
}
