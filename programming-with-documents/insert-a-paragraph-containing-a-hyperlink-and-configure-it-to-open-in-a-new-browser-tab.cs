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

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some introductory text.
        builder.Write("Please visit ");

        // Apply hyperlink styling (blue and underlined).
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;

        // Insert the hyperlink. The method returns a Field object.
        Field field = builder.InsertHyperlink("Aspose", "https://www.aspose.com", false);

        // Cast to FieldHyperlink to set the OpenInNewWindow property.
        if (field is FieldHyperlink hyperlink)
        {
            hyperlink.OpenInNewWindow = true; // Open the link in a new browser tab/window.
        }

        // Reset font formatting to default for the rest of the paragraph.
        builder.Font.ClearFormatting();

        // Complete the paragraph.
        builder.Writeln(" for more info.");

        // Save the document to the local file system.
        doc.Save("HyperlinkParagraph.docx");
    }
}
