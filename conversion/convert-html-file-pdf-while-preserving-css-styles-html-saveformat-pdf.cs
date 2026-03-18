using System;
using Aspose.Words;

class HtmlToPdfConverter
{
    static void Main()
    {
        // Path to the source HTML file.
        string htmlFile = @"C:\Temp\source.html";

        // Path where the resulting PDF will be saved.
        string pdfFile = @"C:\Temp\result.pdf";

        // Load the HTML document. Aspose.Words automatically parses CSS styles,
        // so the visual appearance is preserved when converting.
        Document doc = new Document(htmlFile);

        // Save the loaded document as PDF, preserving the layout and styling.
        doc.Save(pdfFile, SaveFormat.Pdf);
    }
}
