using System;
using Aspose.Words;
using Aspose.Words.Fields;

class DeleteTocExample
{
    static void Main()
    {
        // Path to the source DOCM file that contains a Table of Contents.
        string inputPath = @"C:\Docs\SourceDocument.docm";

        // Load the DOCM document. The constructor automatically detects the format.
        Document doc = new Document(inputPath);

        // Iterate through all fields in the document and remove any TOC (FieldToc) fields.
        // Field.Remove() deletes the field and its result from the document.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
            {
                field.Remove();
            }
        }

        // Path for the output DOC file (macro‑free Word 97‑2007 format).
        string outputPath = @"C:\Docs\ResultDocument.doc";

        // Save the modified document as DOC. The file extension determines the format.
        doc.Save(outputPath, SaveFormat.Doc);
    }
}
