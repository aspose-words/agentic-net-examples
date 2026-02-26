using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class RemoveTocAndSaveAsTemplate
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("input.doc");

        // Iterate through all fields in the document.
        // If a field is a Table of Contents (TOC) field, remove it.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
            {
                // Cast to FieldToc to access the Remove method.
                ((FieldToc)field).Remove();
            }
        }

        // Save the modified document as a Word template (DOT format).
        // The file extension determines the format, but we also specify the format explicitly.
        doc.Save("output.dot", SaveFormat.Dot);
    }
}
