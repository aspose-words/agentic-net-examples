using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsHtmlConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string docxPath = @"C:\Input\SampleDocument.docx";

            // Path where the resulting HTML file will be saved.
            string htmlPath = @"C:\Output\SampleDocument.html";

            // Path for the external CSS file that will be generated.
            string cssPath = @"C:\Output\SampleDocument.css";

            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Configure HTML save options to use an external stylesheet.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export CSS to an external file instead of inline styles.
                CssStyleSheetType = CssStyleSheetType.External,

                // Specify the filename (including full path) for the generated CSS file.
                CssStyleSheetFileName = cssPath,

                // Optional: add a prefix to all generated CSS class names.
                CssClassNamePrefix = "custom_"
            };

            // Save the document as HTML using the configured options.
            doc.Save(htmlPath, saveOptions);
        }
    }
}
