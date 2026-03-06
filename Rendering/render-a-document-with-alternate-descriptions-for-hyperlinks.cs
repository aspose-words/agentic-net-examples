using System.Drawing;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply typical hyperlink formatting.
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;

        // Insert a hyperlink field. The method returns a generic Field object.
        Field field = builder.InsertHyperlink("Visit Aspose", "https://www.aspose.com", false);

        // Cast to FieldHyperlink to set the alternate description (ScreenTip).
        if (field is FieldHyperlink hyperlink)
        {
            hyperlink.ScreenTip = "Open Aspose website in a new browser window";
        }

        // Restore default formatting for subsequent text.
        builder.Font.ClearFormatting();

        // Save the document to disk.
        doc.Save("HyperlinkWithAltDescription.docx");
    }
}
