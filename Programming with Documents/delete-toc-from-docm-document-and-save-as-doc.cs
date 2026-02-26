using System;
using Aspose.Words;
using Aspose.Words.Fields;

class DeleteTocExample
{
    static void Main()
    {
        // Load the macro-enabled DOCM document.
        Document doc = new Document("InputDocument.docm");

        // Iterate through all fields in the document.
        foreach (Field field in doc.Range.Fields)
        {
            // Identify TOC fields.
            if (field.Type == FieldType.FieldTOC)
            {
                // Cast to FieldToc and remove it from the document.
                ((FieldToc)field).Remove();
            }
        }

        // Save the modified document as a legacy DOC file.
        doc.Save("OutputDocument.doc", SaveFormat.Doc);
    }
}
