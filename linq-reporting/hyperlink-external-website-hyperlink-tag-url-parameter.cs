using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text before the hyperlink.
        builder.Write("For more information, visit ");

        // Apply typical hyperlink formatting (blue color, single underline).
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;

        // Insert a hyperlink to an external website.
        // Parameters: display text, URL, isBookmark (false because this is a URL).
        builder.InsertHyperlink("Aspose.Words website", "https://www.aspose.com/words", false);

        // Clear the formatting so subsequent text is not styled as a hyperlink.
        builder.Font.ClearFormatting();
        builder.Writeln(".");

        // Save the document to a file.
        doc.Save("HyperlinkExample.docx");
    }
}
