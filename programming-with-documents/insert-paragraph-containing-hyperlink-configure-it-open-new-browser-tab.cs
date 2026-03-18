using System;
using Aspose.Words;
using Aspose.Words.Fields;

class HyperlinkExample
{
    public static void Run()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text before the hyperlink.
        builder.Write("Visit the ");

        // Insert a hyperlink field. The method returns a generic Field object.
        Field field = builder.InsertHyperlink("Aspose website", "https://www.aspose.com", false);

        // Cast the field to FieldHyperlink to access hyperlink‑specific properties.
        if (field is FieldHyperlink hyperlink)
        {
            // Make the link open in a new browser tab/window.
            hyperlink.OpenInNewWindow = true;

            // Optional: provide a screen tip that appears when the mouse hovers over the link.
            hyperlink.ScreenTip = "Open Aspose website in a new tab";
        }

        // Write the rest of the paragraph and finish it.
        builder.Writeln(" for more information.");

        // Save the document.
        doc.Save("HyperlinkNewWindow.docx");
    }
}

class Program
{
    static void Main(string[] args)
    {
        HyperlinkExample.Run();
    }
}
