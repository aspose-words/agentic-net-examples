using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace RemoveTocExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOC file that contains a table of contents.
            string inputPath = "InputDocument.doc";

            // Path where the resulting DOC file without the TOC will be saved.
            string outputPath = "OutputDocument.doc";

            // Load the existing document.
            Document doc = new Document(inputPath);

            // Iterate over all fields in the document in reverse order.
            // Removing items while iterating forward can skip elements,
            // so we iterate backwards to safely remove TOC fields.
            for (int i = doc.Range.Fields.Count - 1; i >= 0; i--)
            {
                Field field = doc.Range.Fields[i];

                // Identify TOC fields by their type.
                if (field.Type == Aspose.Words.Fields.FieldType.FieldTOC)
                {
                    // Remove the TOC field (the field and its result).
                    field.Remove();
                }
            }

            // Save the modified document in DOC format.
            doc.Save(outputPath, SaveFormat.Doc);
        }
    }
}
