using System;
using Aspose.Words;
using Aspose.Words.Layout;

namespace CommentPdfConversion
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample DOC file with a comment.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a paragraph of text.
            builder.Writeln("This is a sample paragraph that will have a comment.");

            // Create a comment, set its metadata, and add text to it.
            Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
            comment.SetText("This is a comment attached to the paragraph.");

            // Attach the comment to the current paragraph.
            builder.CurrentParagraph?.AppendChild(comment);

            // Save the DOC file.
            const string docPath = "sample.doc";
            doc.Save(docPath);

            // Load the DOC file for conversion.
            Document loadedDoc = new Document(docPath);

            // Configure the layout to render comments as PDF annotations.
            loadedDoc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;

            // Rebuild the layout after changing the option.
            loadedDoc.UpdatePageLayout();

            // Save the document as PDF; comments will be visible.
            const string pdfPath = "sample.pdf";
            loadedDoc.Save(pdfPath);
        }
    }
}
