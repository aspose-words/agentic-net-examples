using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string sourcePath = @"C:\Docs\Source.doc";

        // Path where the resulting TIFF will be saved.
        string resultPath = @"C:\Docs\Result.tiff";

        // Load the existing document.
        Document doc = new Document(sourcePath);

        // Iterate through all fields in the document and remove any Table of Contents fields.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                field.Remove();
        }

        // Re‑calculate layout after removing the TOC.
        doc.UpdatePageLayout();

        // Prepare image save options for TIFF output.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Save the modified document as a multi‑page TIFF image.
        doc.Save(resultPath, saveOptions);
    }
}
