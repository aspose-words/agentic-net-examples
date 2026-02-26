using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertWithDocxSaveOptions
{
    static void Main()
    {
        // Path to the source document.
        string inputPath = @"MyDir\SourceDocument.docx";

        // Load the document.
        Document doc = new Document(inputPath);

        // Create a save options object suitable for DOCX format.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();

        // Configure the save options – for example, enforce strict OOXML compliance.
        saveOptions.SaveFormat = SaveFormat.Docx;
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Path to the output document.
        string outputPath = @"ArtifactsDir\ConvertedDocument.docx";

        // Save the document using the configured options.
        doc.Save(outputPath, saveOptions);
    }
}
