using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ContentControlToHtml
{
    class Program
    {
        static void Main()
        {
            // Path to the source Word document that contains content controls (StructuredDocumentTag nodes).
            string inputPath = @"C:\Docs\InputWithContentControls.docx";

            // Path where the resulting HTML file will be saved.
            string outputPath = @"C:\Docs\Output.html";

            // Load the Word document.
            Document doc = new Document(inputPath);

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Preserve round‑trip information so that content control attributes are exported
                // as custom data‑attributes (e.g., data-tag, data-title) in the HTML.
                ExportRoundtripInformation = true,

                // Keep form fields as interactive HTML elements rather than plain text.
                ExportTextInputFormFieldAsText = false,
                ExportDropDownFormFieldAsText = false,

                // Optional: make the output more readable.
                PrettyFormat = true
            };

            // Save the document as HTML using the configured options.
            doc.Save(outputPath, htmlOptions);
        }
    }
}
