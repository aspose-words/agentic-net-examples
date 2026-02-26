using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDocmWithOptions
{
    static void Main()
    {
        // Paths to the source and destination files.
        string dataDir = @"C:\Data\";
        string inputPath = Path.Combine(dataDir, "input.docx");
        string outputPath = Path.Combine(dataDir, "output.docm");

        // Load the source document.
        Document doc = new Document(inputPath);

        // Create OoxmlSaveOptions for the DOCM format.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm)
        {
            // Example advanced settings:
            UpdateFields = true,                     // Update fields before saving.
            Password = "Secret",                     // Encrypt the saved document.
            Compliance = OoxmlCompliance.Iso29500_2008_Strict, // Enforce strict OOXML compliance.
            UseHighQualityRendering = true           // Use high‑quality rendering algorithms.
        };

        // Save the document as a macro‑enabled DOCM file using the specified options.
        doc.Save(outputPath, saveOptions);
    }
}
