using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class MhtmlToPdfConverter
{
    static void Main()
    {
        // Path to the source MHTML file.
        const string inputMhtmlPath = @"C:\Docs\source.mht";

        // Path where the resulting PDF will be saved.
        const string outputPdfPath = @"C:\Docs\result.pdf";

        // Load options for HTML/MHTML files.
        // Setting the LoadFormat explicitly ensures the file is treated as MHTML.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            LoadFormat = LoadFormat.Mhtml
        };

        // Load the MHTML document into an Aspose.Words Document object.
        Document document = new Document(inputMhtmlPath, loadOptions);

        // Save the document as PDF. The SaveFormat.Pdf forces PDF output
        // and preserves all embedded images, styles, fonts, etc.
        document.Save(outputPdfPath, SaveFormat.Pdf);
    }
}
