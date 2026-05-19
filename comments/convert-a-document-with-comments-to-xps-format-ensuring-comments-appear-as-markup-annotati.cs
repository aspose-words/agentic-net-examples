using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain a comment.
        builder.Writeln("This is a paragraph that will have a comment attached to it.");

        // Create a comment with author metadata.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("This is a sample comment displayed as an annotation.");

        // Anchor the comment to the whole paragraph.
        // The comment must be linked with a CommentRangeStart and CommentRangeEnd that share the same Id.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        paragraph.PrependChild(new CommentRangeStart(doc, comment.Id));
        paragraph.AppendChild(new CommentRangeEnd(doc, comment.Id));
        paragraph.AppendChild(comment);

        // Ensure comments are rendered as annotations (balloons) in the output.
        // For XPS the comments are shown as markup annotations by default.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInBalloons;

        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document to XPS format with comments preserved.
        string xpsPath = Path.Combine(outputDir, "DocumentWithComments.xps");
        XpsSaveOptions saveOptions = new XpsSaveOptions();
        doc.Save(xpsPath, saveOptions);
    }
}
