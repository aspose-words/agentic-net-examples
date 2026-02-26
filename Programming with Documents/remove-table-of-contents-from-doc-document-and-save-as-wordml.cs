using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;
using System.Collections.Generic;

class RemoveTocAndSaveAsWordml
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("input.doc");

        // Collect all TOC fields first – we cannot modify the collection while iterating it.
        List<Field> tocFields = new List<Field>();
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                tocFields.Add(field);
        }

        // Remove the collected TOC fields.
        foreach (Field field in tocFields)
            field.Remove();

        // Save the modified document in WordML (Word 2003 XML) format.
        doc.Save("output.xml", SaveFormat.WordML);
    }
}
