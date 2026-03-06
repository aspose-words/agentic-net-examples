using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfTemplateToPs
{
    static void Main()
    {
        // Load the PDF template that contains fields or expressions.
        Document doc = new Document("Template.pdf");

        // Ensure all fields (including expressions) are evaluated before saving.
        doc.UpdateFields();

        // Configure PostScript save options.
        PsSaveOptions psOptions = new PsSaveOptions
        {
            // Explicitly set the format to PostScript.
            SaveFormat = SaveFormat.Ps
        };

        // Save the evaluated document as a PostScript file.
        doc.Save("Result.ps", psOptions);
    }
}
