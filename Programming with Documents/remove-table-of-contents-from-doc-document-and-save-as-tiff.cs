using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields; // Added for Field and FieldType

class RemoveTocAndConvertToTiff
{
    static void Main()
    {
        // Path to the source DOC document.
        const string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path where the resulting TIFF image will be saved.
        const string outputPath = @"C:\Docs\ResultImage.tiff";

        // Load the existing Word document.
        Document doc = new Document(inputPath);

        // Remove all Table of Contents (TOC) fields.
        // Iterate backwards to avoid collection modification issues.
        for (int i = doc.Range.Fields.Count - 1; i >= 0; i--)
        {
            Field field = doc.Range.Fields[i];
            if (field.Type == FieldType.FieldTOC)
                field.Remove();
        }

        // Prepare image save options for TIFF format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Save the modified document as a TIFF image.
        doc.Save(outputPath, saveOptions);
    }
}
