using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (any supported format, e.g., DOCX).
        Document doc = new Document("Input.docx");

        // Configure save options for the DOCM (macro‑enabled Word) format.
        // The constructor that accepts a SaveFormat ensures the correct format is used.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm);

        // Optional: set the OOXML compliance level if specific features are required.
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

        // Save the document as a DOCM file using the configured options.
        doc.Save("Output.docm", saveOptions);
    }
}
