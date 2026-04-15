using System;
using System.IO;
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

        // Add some introductory text.
        builder.Writeln("Please click the link below:");

        // Set formatting for the hyperlink text.
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;

        // Insert the hyperlink.
        Field field = builder.InsertHyperlink("Open Aspose website", "https://www.aspose.com", false);

        // Configure the hyperlink to open in a new browser tab/window.
        if (field is FieldHyperlink hyperlink)
        {
            hyperlink.OpenInNewWindow = true;
        }

        // Reset formatting after the hyperlink.
        builder.Font.ClearFormatting();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Hyperlink.docx");
        doc.Save(outputPath);
    }
}
