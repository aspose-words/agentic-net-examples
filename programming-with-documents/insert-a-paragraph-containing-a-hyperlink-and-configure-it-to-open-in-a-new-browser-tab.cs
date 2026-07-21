using System;
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

        // Write introductory text.
        builder.Write("Visit the ");

        // Apply hyperlink formatting.
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;

        // Insert the hyperlink.
        Field field = builder.InsertHyperlink("Aspose.Words site", "https://www.aspose.com/words", false);

        // Set the hyperlink to open in a new browser tab/window.
        if (field is FieldHyperlink hyperlink)
        {
            hyperlink.OpenInNewWindow = true;
        }

        // Reset formatting to default.
        builder.Font.ClearFormatting();

        // Complete the sentence.
        builder.Writeln(" for more information.");

        // Save the document.
        doc.Save("HyperlinkNewTab.docx");
    }
}
