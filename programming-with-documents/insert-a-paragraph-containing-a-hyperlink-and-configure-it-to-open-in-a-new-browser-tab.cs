using System.Drawing;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some introductory text.
        builder.Write("Please visit ");

        // Apply typical hyperlink formatting.
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;

        // Insert the hyperlink. The method returns a Field object.
        Field field = builder.InsertHyperlink("Aspose website", "https://www.aspose.com", false);

        // Cast to FieldHyperlink to enable opening in a new browser tab/window.
        if (field is FieldHyperlink hyperlink)
        {
            hyperlink.OpenInNewWindow = true;
        }

        // Reset formatting to default for subsequent text.
        builder.Font.ClearFormatting();

        // Finish the paragraph.
        builder.Writeln(" for more information.");

        // Save the document to the local file system.
        doc.Save("HyperlinkNewTab.docx");
    }
}
