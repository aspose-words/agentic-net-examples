using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace DeleteTocExample
{
    class Program
    {
        static void Main()
        {
            // Load the macro‑enabled Word document (DOCM).
            // The Document constructor handles opening the file and detecting its format.
            Document doc = new Document("InputDocument.docm");

            // Iterate over all fields in the document.
            // FieldType.FieldTOC identifies a Table of Contents field.
            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type == FieldType.FieldTOC)
                {
                    // Remove the TOC field from the document.
                    // The Remove method returns the node that follows the removed field,
                    // but we do not need the return value here.
                    field.Remove();
                }
            }

            // Save the modified document as a Word template (DOT).
            // The Save method automatically selects the format based on the file extension.
            doc.Save("OutputTemplate.dot");
        }
    }
}
