using System;
using Aspose.Words;
using Aspose.Words.Fields;

class TocRemover
{
    static void Main()
    {
        // Load the macro‑enabled Word document (DOCM).
        Document doc = new Document("InputDocument.docm");

        // Iterate through all fields in the document and remove any Table of Contents fields.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
            {
                // Remove the entire TOC field, including its result text.
                field.Remove();
            }
        }

        // Save the modified document as a macro‑enabled template (DOTM).
        doc.Save("OutputTemplate.dotm");
    }
}
