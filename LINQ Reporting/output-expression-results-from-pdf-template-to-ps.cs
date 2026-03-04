using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace OutputExpressionResultsFromPdfTemplateToPs
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the PDF template (the constructor with a file path loads the document)
            Document doc = new Document("Template.pdf");

            // Ensure that all fields (including expressions) are evaluated before saving
            doc.UpdateFields();

            // Configure PostScript save options
            PsSaveOptions psOptions = new PsSaveOptions
            {
                // Explicitly set the format to PostScript (optional, PsSaveOptions already implies this)
                SaveFormat = SaveFormat.Ps,
                // No booklet layout is required for a simple conversion
                UseBookFoldPrintingSettings = false
            };

            // Save the evaluated document as a PostScript file
            doc.Save("Result.ps", psOptions);
        }
    }
}
