using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToPsConverter
{
    static void Main()
    {
        // Load the PDF template into an Aspose.Words Document.
        Document doc = new Document("Template.pdf");

        // Create a PsSaveOptions object to control the conversion to PostScript.
        PsSaveOptions psOptions = new PsSaveOptions
        {
            // Explicitly set the format to PostScript (Ps).
            SaveFormat = SaveFormat.Ps,

            // Example option: do not use booklet layout.
            UseBookFoldPrintingSettings = false
        };

        // Save the document as a PostScript file using the specified options.
        doc.Save("Output.ps", psOptions);
    }
}
