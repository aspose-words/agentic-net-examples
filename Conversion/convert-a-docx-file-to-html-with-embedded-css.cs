using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocxToHtmlConverter
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputFile = "input.docx";

        // Path where the resulting HTML file will be saved.
        string outputFile = "output.html";

        // Load the DOCX document.
        Document doc = new Document(inputFile);

        // Configure HTML save options to embed CSS within the HTML file.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            CssStyleSheetType = CssStyleSheetType.Embedded   // CSS will be placed inside a <style> tag.
        };

        // Save the document as HTML using the configured options.
        doc.Save(outputFile, saveOptions);
    }
}
