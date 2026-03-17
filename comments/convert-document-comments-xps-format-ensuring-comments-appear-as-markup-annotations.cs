using System;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document that contains comments.
        Document doc = new Document("Comments.docx");

        // Render comments as markup annotations.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;

        // Rebuild the layout after changing the display mode.
        doc.UpdatePageLayout();

        // Create XPS save options (default settings are sufficient).
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        // Save the document to XPS format; comments will appear as annotations.
        doc.Save("Comments.xps", xpsOptions);
    }
}
