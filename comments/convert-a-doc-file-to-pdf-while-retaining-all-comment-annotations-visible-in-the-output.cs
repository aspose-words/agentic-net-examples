using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for the example files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a sample DOC file with a comment.
        string docPath = Path.Combine(outputDir, "SampleWithComments.doc");
        CreateSampleDocumentWithComment(docPath);

        // Load the DOC file.
        Document doc = new Document(docPath);

        // Configure the layout to render comments as PDF annotations.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;

        // Rebuild the layout after changing the option.
        doc.UpdatePageLayout();

        // Save the document as PDF. Comments will appear as visible annotations.
        string pdfPath = Path.Combine(outputDir, "SampleWithComments.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);
    }

    private static void CreateSampleDocumentWithComment(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text.
        builder.Writeln("This is a paragraph that will have a comment.");

        // Create a comment and attach it to the current paragraph.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("Review this paragraph for clarity.");
        builder.CurrentParagraph.AppendChild(comment);

        // Save the DOC file.
        doc.Save(filePath);
    }
}
