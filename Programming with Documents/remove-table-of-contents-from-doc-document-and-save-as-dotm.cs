using System;
using Aspose.Words;
using Aspose.Words.Fields;

class RemoveTocAndSaveAsDotm
{
    static void Main()
    {
        // Load the existing DOC file.
        Document doc = new Document("InputDocument.doc");

        // Iterate through all fields in the document.
        // Remove any field that is a Table of Contents (TOC) field.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                field.Remove();
        }

        // Save the modified document as a macro‑enabled template (DOTM).
        doc.Save("OutputDocument.dotm", SaveFormat.Dotm);
    }
}
