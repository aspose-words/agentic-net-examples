using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an MHTML document. HtmlLoadOptions can be used for HTML/MHTML loading.
        var loadOptions = new HtmlLoadOptions
        {
            // Enable VML support if the source contains VML images.
            SupportVml = true
        };

        // Load the document from an MHTML file.
        Document doc = new Document("Input.mht", loadOptions);

        // Example of contextual object member access: retrieve the text of the first paragraph.
        string firstParagraph = doc.FirstSection.Body.Paragraphs[0].GetText();
        Console.WriteLine("First paragraph: " + firstParagraph);

        // Configure save options to export the document as MHTML using CID URLs for resources.
        var saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportCidUrlsForMhtmlResources = true, // Use CID URLs.
            ExportFontResources = true,            // Include font resources.
            PrettyFormat = true                    // Make the output more readable.
        };

        // Save the document as MHTML with the specified options.
        doc.Save("Output.mht", saveOptions);
    }
}
