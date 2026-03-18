using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a line that will serve as a comment and insert a hyperlink to an external resource.
        builder.Writeln("Comment:");
        builder.Font.Color = Color.Blue;               // Make the hyperlink appear like a typical web link.
        builder.Font.Underline = Underline.Single;
        builder.InsertHyperlink(
            "Aspose.Words Documentation",               // Text displayed in the document.
            "https://docs.aspose.com/words/net/",       // URL of the external resource.
            false);                                     // Not a bookmark; it's a URL.
        builder.Font.ClearFormatting();                // Reset formatting for subsequent text.

        // Configure PDF save options so that hyperlinks are saved normally (clickable in the PDF).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OpenHyperlinksInNewWindow = false // Links open in the same window/tab; required for most PDF readers.
        };

        // Save the document as a PDF file.
        doc.Save("CommentWithHyperlink.pdf", pdfOptions);
    }
}
