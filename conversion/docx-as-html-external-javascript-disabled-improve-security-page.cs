using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("InputDocument.docx");

        // Configure HTML save options.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Disable JavaScript in links for security.
        htmlOptions.RemoveJavaScriptFromLinks = true;

        // Save the document as HTML using the configured options.
        doc.Save("OutputDocument.html", htmlOptions);
    }
}
