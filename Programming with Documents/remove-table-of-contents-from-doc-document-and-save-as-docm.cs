using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

class RemoveTocAndSaveAsDocm
{
    static void Main()
    {
        // Path to the source DOC file.
        const string inputFile = @"C:\Docs\SourceDocument.doc";

        // Path where the resulting DOCM file will be saved.
        const string outputFile = @"C:\Docs\ResultDocument.docm";

        // Load the existing document (lifecycle rule: load).
        Document doc = new Document(inputFile);

        // Collect all Table of Contents (TOC) fields in the document.
        List<Field> tocFields = new List<Field>();
        foreach (Field field in doc.Range.Fields)
        {
            // FieldType.FieldTOC identifies a TOC field.
            if (field.Type == FieldType.FieldTOC)
                tocFields.Add(field);
        }

        // Remove each TOC field from the document.
        foreach (Field toc in tocFields)
        {
            // The Remove method deletes the field node and its result.
            toc.Remove();
        }

        // Optionally update remaining fields after removal.
        doc.UpdateFields();

        // Save the modified document as a macro‑enabled DOCM file (lifecycle rule: save).
        doc.Save(outputFile, SaveFormat.Docm);
    }
}
