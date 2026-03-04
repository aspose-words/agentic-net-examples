using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDocm
{
    static void Main()
    {
        // Load the source document (any format supported by Aspose.Words).
        Document doc = new Document("Input.docx");

        // For saving as a macro‑enabled Word document (DOCM) we must use OoxmlSaveOptions.
        // HtmlSaveOptions, MhtmlSaveOptions, and EpubSaveOptions are intended for HTML‑based formats
        // and cannot be used to produce a DOCM file.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm);

        // Optional: configure additional options for the DOCM output.
        saveOptions.Compliance = OoxmlCompliance.Ecma376_2006; // OOXML version.
        saveOptions.PrettyFormat = true;                     // Human‑readable XML.

        // Save the document as a DOCM file using the configured options.
        doc.Save("Output.docm", saveOptions);
    }
}
