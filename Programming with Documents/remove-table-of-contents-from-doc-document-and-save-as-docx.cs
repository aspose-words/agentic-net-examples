using System;
using Aspose.Words;
using Aspose.Words.Fields; // Needed for Field and FieldType

class RemoveTocExample
{
    static void Main()
    {
        // Load the existing DOC file.
        Document doc = new Document("InputDocument.doc");

        // Collect all TOC fields first because removing while iterating the collection can cause enumeration issues.
        var tocFields = new System.Collections.Generic.List<Field>();
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                tocFields.Add(field);
        }

        // Remove the collected TOC fields.
        foreach (Field toc in tocFields)
        {
            toc.Remove();
        }

        // Save the modified document as DOCX.
        doc.Save("OutputDocument.docx");
    }
}
