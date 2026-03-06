using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields; // Added for Field and FieldType
using Aspose.Words.Saving;

class RemoveTocAndSaveAsPs
{
    static void Main()
    {
        // Paths to the source DOC file and the destination PS file.
        string inputPath = "input.doc";
        string outputPath = "output.ps";

        // Load the existing Word document.
        Document doc = new Document(inputPath);

        // Collect all Table of Contents (TOC) fields in the document.
        List<Field> tocFields = new List<Field>();
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                tocFields.Add(field);
        }

        // Remove each TOC field together with its result.
        foreach (Field toc in tocFields)
        {
            toc.Remove();
        }

        // Prepare PostScript save options.
        PsSaveOptions saveOptions = new PsSaveOptions
        {
            SaveFormat = SaveFormat.Ps
        };

        // Save the modified document as a PostScript file.
        doc.Save(outputPath, saveOptions);
    }
}
