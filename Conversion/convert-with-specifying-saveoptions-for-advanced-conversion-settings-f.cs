using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToDocm
{
    static void Main()
    {
        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, DOCM!");

        // Create OoxmlSaveOptions for the DOCM format.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm);

        // Advanced conversion settings (examples):
        // Enforce strict OOXML compliance.
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Strict;
        // Encrypt the saved file with a password.
        saveOptions.Password = "Secret123";
        // Ensure fields are updated before saving.
        saveOptions.UpdateFields = true;
        // Use high‑quality rendering algorithms.
        saveOptions.UseHighQualityRendering = true;

        // Save the document as a macro‑enabled DOCM file using the specified options.
        doc.Save("ConvertedDocument.docm", saveOptions);
    }
}
