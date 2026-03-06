using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace RemoveTocExample
{
    class Program
    {
        static void Main()
        {
            // Load the existing DOC document.
            Document doc = new Document("InputDocument.doc");

            // Iterate through all fields in the document.
            // If a field is a Table of Contents (FieldToc), remove it.
            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type == FieldType.FieldTOC)
                {
                    // The Remove method returns the node after the removed field,
                    // but we do not need the return value here.
                    field.Remove();
                }
            }

            // Save the modified document as a Word template (DOT format).
            doc.Save("OutputTemplate.dot", SaveFormat.Dot);
        }
    }
}
