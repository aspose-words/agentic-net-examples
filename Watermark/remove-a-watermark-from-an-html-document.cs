using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveWatermarkFromHtml
{
    static void Main()
    {
        // Load the HTML document from file.
        Document doc = new Document("input.html");

        // Remove the watermark if one is present.
        if (doc.Watermark.Type != WatermarkType.None)
        {
            doc.Watermark.Remove();
        }

        // Save the modified document back to HTML format.
        doc.Save("output.html", SaveFormat.Html);
    }
}
