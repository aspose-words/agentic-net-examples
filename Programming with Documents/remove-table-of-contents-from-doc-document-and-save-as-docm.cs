using System;
using Aspose.Words;
using Aspose.Words.Fields;

class RemoveTocExample
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path for the resulting DOCM file (macro‑enabled format).
        string outputPath = @"C:\Docs\SourceDocument_NoToc.docm";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Iterate through all fields in the document and remove any Table of Contents fields.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                field.Remove();
        }

        // Save the modified document as a DOCM file.
        doc.Save(outputPath);
    }
}
