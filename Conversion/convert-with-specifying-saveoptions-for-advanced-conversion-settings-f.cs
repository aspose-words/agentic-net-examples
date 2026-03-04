using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (replace with your actual file path).
        Document doc = new Document("Input.docx");

        // Create OoxmlSaveOptions for the DOCM format.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm);

        // Advanced conversion settings (customize as needed).
        saveOptions.Password = "MySecretPassword";          // Encrypt the saved DOCM.
        saveOptions.UpdateFields = true;                    // Update fields before saving.
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Strict; // Enforce strict OOXML compliance.
        saveOptions.UseHighQualityRendering = true;        // Enable high‑quality rendering.

        // Save the document as a macro‑enabled DOCM using the specified options.
        doc.Save("Output.docm", saveOptions);
    }
}
