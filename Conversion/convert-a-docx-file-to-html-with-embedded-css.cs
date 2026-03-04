using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputFile = @"C:\Docs\Sample.docx";

        // Path where the resulting HTML file will be saved.
        string outputFile = @"C:\Docs\Sample.html";

        // Load the DOCX document.
        Document doc = new Document(inputFile);

        // Configure HTML save options to embed CSS directly into the HTML file.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            CssStyleSheetType = CssStyleSheetType.Embedded // CSS will be placed inside a <style> element.
        };

        // Save the document as HTML with the specified options.
        doc.Save(outputFile, saveOptions);
    }
}
