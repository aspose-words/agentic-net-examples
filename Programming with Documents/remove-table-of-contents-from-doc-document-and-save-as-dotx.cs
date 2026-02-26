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
        // If a field is a Table of Contents (FieldToc), remove it.
        foreach (Field field in doc.Range.Fields)
        {
            if (field is FieldToc toc)
            {
                toc.Remove();
            }
        }

        // Save the modified document as a DOTX template.
        // The file extension determines the SaveFormat (DOTX = macro‑free template).
        doc.Save("OutputDocument.dotx");
    }
}
