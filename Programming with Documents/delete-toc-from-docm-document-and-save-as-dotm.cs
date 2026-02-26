using System;
using Aspose.Words;
using Aspose.Words.Fields;

class DeleteTocAndSaveAsTemplate
{
    static void Main()
    {
        // Load the existing macro‑enabled document (DOCM).
        Document doc = new Document("InputDocument.docm");

        // Iterate over all fields in the document and remove those that are TOC fields.
        // Field.Type == FieldType.FieldTOC identifies a Table of Contents field.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                field.Remove();
        }

        // Save the modified document as a macro‑enabled template (DOTM).
        // The Save method infers the format from the file extension.
        doc.Save("OutputTemplate.dotm");
    }
}
