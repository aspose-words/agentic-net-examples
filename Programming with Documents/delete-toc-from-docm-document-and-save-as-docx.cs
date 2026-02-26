using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the macro‑enabled DOCM file.
        Document doc = new Document("input.docm");

        // Iterate through all fields in the document and remove those that are TOC fields.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                field.Remove(); // Deletes the TOC field from the document.
        }

        // Save the modified document as a regular DOCX file.
        doc.Save("output.docx");
    }
}
