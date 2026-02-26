using System;
using Aspose.Words;
using Aspose.Words.Fields;

class RemoveTocAndSaveAsTemplate
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // Iterate through all fields in the document.
        // Remove any field that is a Table of Contents (TOC) field.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
            {
                // The Remove method returns the node that follows the removed field,
                // but we do not need the return value here.
                field.Remove();
            }
        }

        // Optionally update remaining fields (e.g., page numbers) after removal.
        doc.UpdateFields();

        // Save the modified document as a macro‑enabled template (DOTM).
        // The file extension determines the format, so no explicit SaveFormat is required.
        doc.Save("OutputDocument.dotm");
    }
}
