using System;
using System.Drawing;
using Aspose.Words;

class Program
{
    static void Main(string[] args)
    {
        HyperlinkExample.Run();
    }
}

class HyperlinkExample
{
    public static void Run()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Data fields that provide the URL and the text to display.
        string url = "https://www.example.com";
        string displayText = "Visit Example";

        // Write some introductory text before the hyperlink.
        builder.Write("For more information, please ");

        // Apply typical hyperlink formatting (blue and underlined).
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;

        // Insert the hyperlink. The third argument is false because we are linking to a URL,
        // not a bookmark inside the document.
        builder.InsertHyperlink(displayText, url, false);

        // Clear the formatting so subsequent text uses the default style.
        builder.Font.ClearFormatting();

        // Finish the sentence.
        builder.Writeln(".");

        // Save the document to disk.
        doc.Save("Hyperlink.docx");
    }
}
