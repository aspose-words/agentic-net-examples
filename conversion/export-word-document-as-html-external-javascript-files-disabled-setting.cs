using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("input.docx");

        // Create HTML save options and disable JavaScript in links.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.RemoveJavaScriptFromLinks = true; // removes any "javascript:" hrefs.

        // Save the document as HTML using the configured options.
        doc.Save("output.html", htmlOptions);
    }
}
