using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class RemoveTocAndSaveAsTxt
{
    static void Main()
    {
        // Load the source DOC document.
        // (Assumes the file "Input.doc" exists in the same folder as the executable.)
        Document doc = new Document("Input.doc");

        // Collect all Table of Contents (TOC) fields in the document.
        // Field.Type == FieldType.FieldTOC identifies a TOC field.
        List<Field> tocFields = new List<Field>();
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                tocFields.Add(field);
        }

        // Remove each TOC field from the document.
        // Removing the field node also removes its result text.
        foreach (Field toc in tocFields)
        {
            toc.Remove();
        }

        // Save the modified document as plain text.
        TxtSaveOptions txtOptions = new TxtSaveOptions();
        txtOptions.SaveFormat = SaveFormat.Text; // Explicitly set the format.
        doc.Save("Output.txt", txtOptions);
    }
}
