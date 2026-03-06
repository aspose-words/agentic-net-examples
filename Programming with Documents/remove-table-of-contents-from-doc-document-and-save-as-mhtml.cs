using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields; // Added namespace for Field and FieldType

class RemoveTocAndSaveAsMhtml
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // Iterate through all fields in the document and remove any Table of Contents fields.
        // FieldType.FieldTOC identifies a TOC field.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                field.Remove();
        }

        // Save the modified document as MHTML.
        // The SaveFormat enumeration value Mhtml tells Aspose.Words to output MHTML.
        doc.Save("OutputDocument.mhtml", SaveFormat.Mhtml);
    }
}
