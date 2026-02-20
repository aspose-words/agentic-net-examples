using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a hyperlink with custom formatting.
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;
        builder.InsertHyperlink("Aspose", "https://www.aspose.com", false);
        builder.Font.ClearFormatting();

        // Retrieve the inserted hyperlink field and set an alternate description (ScreenTip).
        FieldHyperlink hyperlink = (FieldHyperlink)doc.Range.Fields[doc.Range.Fields.Count - 1];
        hyperlink.ScreenTip = "Visit Aspose website";

        // Save the document as Markdown using reference‑style links.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        saveOptions.LinkExportMode = MarkdownLinkExportMode.Reference;
        doc.Save("Hyperlinks.md", saveOptions);
    }
}
