using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DotToHtmlConverter
{
    static void Main()
    {
        // Path to the folder that contains the source .dot file and where the HTML will be saved.
        string folderPath = @"C:\Documents\";

        // Load the DOT template. The LoadFormat is detected automatically.
        Document doc = new Document(System.IO.Path.Combine(folderPath, "template.dot"));

        // Configure HTML save options. Use HTML5 standard for the output.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            HtmlVersion = HtmlVersion.Html5,
            PrettyFormat = true   // Optional: make the HTML output more readable.
        };

        // Save the document as an HTML file.
        doc.Save(System.IO.Path.Combine(folderPath, "output.html"), htmlOptions);
    }
}
