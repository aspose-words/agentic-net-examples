using System;
using Aspose.Words;
using Aspose.Words.Fields;

class DeleteTocAndSaveAsTemplate
{
    static void Main()
    {
        // Path to the source DOCM file that contains a Table of Contents.
        string sourcePath = @"C:\Docs\SourceDocument.docm";

        // Load the DOCM document using the provided Document(string) constructor.
        Document doc = new Document(sourcePath);

        // Iterate over all fields in the document.
        // The Field.Type property identifies the kind of field.
        // Remove every field whose type is TOC (Table of Contents).
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
            {
                // The Remove method deletes the field from the document.
                field.Remove();
            }
        }

        // Path for the resulting template file. The .dot extension tells Aspose.Words
        // to save the document in the Microsoft Word Template format.
        string outputPath = @"C:\Docs\ResultTemplate.dot";

        // Save the modified document using the provided Save(string) method.
        // The format is inferred from the file extension (.dot → SaveFormat.Dot).
        doc.Save(outputPath);
    }
}
