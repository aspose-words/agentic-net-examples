using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string inputDocxPath = @"C:\Docs\SourceDocument.docx";

            // Path where the resulting HTML file will be saved.
            string outputHtmlPath = @"C:\Docs\ResultDocument.html";

            // Path for the external CSS file that will be generated.
            string outputCssPath = @"C:\Docs\ResultStyles.css";

            // Load the DOCX document.
            Document doc = new Document(inputDocxPath);

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export CSS to an external stylesheet.
                CssStyleSheetType = CssStyleSheetType.External,

                // Specify the filename for the external CSS file.
                CssStyleSheetFileName = outputCssPath,

                // Optional: add a prefix to all generated CSS class names.
                CssClassNamePrefix = "custom_"
            };

            // Save the document as HTML using the configured options.
            doc.Save(outputHtmlPath, htmlOptions);
        }
    }
}
