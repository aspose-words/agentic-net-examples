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

        // Write some introductory text.
        builder.Write("Please click the following link: ");

        // Apply hyperlink formatting (blue color, single underline).
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;

        // Insert the hyperlink. The method returns a Field object.
        Field field = builder.InsertHyperlink("Aspose.Words", "https://www.aspose.com/words", false);

        // Cast to FieldHyperlink to configure it to open in a new browser tab/window.
        if (field is FieldHyperlink hyperlink)
        {
            hyperlink.OpenInNewWindow = true;
        }

        // Reset font formatting to default for subsequent text.
        builder.Font.ClearFormatting();

        // End the paragraph.
        builder.Writeln();

        // Save the document to the local file system.
        doc.Save("HyperlinkParagraph.docx");
    }
}
