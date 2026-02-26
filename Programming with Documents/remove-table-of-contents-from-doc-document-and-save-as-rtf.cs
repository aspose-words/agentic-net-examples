using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields; // Added for Field and FieldType
using Aspose.Words.Saving;

class RemoveTocAndSaveAsRtf
{
    static void Main()
    {
        // Path to the source DOC document.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path where the resulting RTF will be saved.
        string outputPath = @"C:\Docs\ResultDocument.rtf";

        // Load the DOC document.
        Document doc = new Document(inputPath);

        // Collect all Table of Contents (TOC) fields.
        List<Field> tocFields = new List<Field>();
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                tocFields.Add(field);
        }

        // Remove each TOC field from the document.
        foreach (Field toc in tocFields)
        {
            // Remove the field and its associated result.
            toc.Remove();
        }

        // Update remaining fields (e.g., page numbers) after removal.
        doc.UpdateFields();

        // Save the modified document as RTF using RtfSaveOptions.
        RtfSaveOptions rtfOptions = new RtfSaveOptions
        {
            // Example option: make the RTF file smaller (set to true if RTL text is not needed).
            ExportCompactSize = false,
            ExportImagesForOldReaders = true
        };

        doc.Save(outputPath, rtfOptions);
    }
}
