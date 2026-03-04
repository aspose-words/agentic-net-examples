using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (any supported format).
        Document doc = new Document("Input.docx");

        // Create a SaveOptions object suitable for DOCX format.
        SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Docx);

        // The created object is an OoxmlSaveOptions; configure DOCX‑specific settings.
        if (saveOptions is OoxmlSaveOptions ooxmlOptions)
        {
            // Example: enforce strict ISO/IEC 29500:2008 OOXML compliance.
            ooxmlOptions.Compliance = OoxmlCompliance.Iso29500_2008_Strict;
        }

        // Save the document as DOCX using the configured options.
        doc.Save("Output.docx", saveOptions);
    }
}
