using System;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with a comment.");

        // Add a comment to the last paragraph.
        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.AppendChild(new Run(doc, "This is a comment added programmatically."));
        comment.Paragraphs.Add(commentParagraph);
        builder.CurrentParagraph.AppendChild(comment);

        // Render comments as markup annotations.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;

        // Rebuild the layout after changing the display mode.
        doc.UpdatePageLayout();

        // Save the document to XPS format; comments will appear as annotations.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        doc.Save("Comments.xps", xpsOptions);
    }
}
