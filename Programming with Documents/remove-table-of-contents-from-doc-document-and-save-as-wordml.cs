using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class RemoveTocAndSaveAsWordml
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("input.doc");

        // Iterate through all fields in the document.
        // If a field is a Table of Contents (FieldToc), remove it.
        foreach (Field field in doc.Range.Fields)
        {
            if (field is FieldToc tocField)
            {
                // Remove the TOC field from the document.
                tocField.Remove();
            }
        }

        // Save the modified document in WordML (Word 2003 XML) format.
        // Using the overload that specifies the SaveFormat directly.
        doc.Save("output.xml", SaveFormat.WordML);
    }
}
