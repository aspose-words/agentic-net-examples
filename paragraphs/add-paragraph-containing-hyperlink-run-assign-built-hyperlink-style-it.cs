using System.Drawing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text before the hyperlink.
        builder.Write("Visit the ");

        // Apply the built‑in Hyperlink character style to the next run.
        builder.Font.StyleIdentifier = StyleIdentifier.Hyperlink;
        // Ensure typical hyperlink appearance (optional).
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;

        // Insert the hyperlink field.
        builder.InsertHyperlink("Aspose website", "https://www.aspose.com", false);

        // Reset formatting for any following text.
        builder.Font.ClearFormatting();

        // Write the rest of the paragraph and finish it.
        builder.Writeln(" for more information.");

        // Save the document.
        doc.Save("HyperlinkParagraph.docx");
    }
}
