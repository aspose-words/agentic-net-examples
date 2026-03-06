using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Path to the source DOCM file that contains a Table of Contents.
        string sourcePath = @"C:\Input\DocumentWithToc.docm";

        // Load the DOCM document.
        Document doc = new Document(sourcePath);

        // Iterate over all fields in the document and remove those that are TOC fields.
        // Field.Type == FieldType.FieldTOC identifies a Table of Contents field.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
            {
                // Remove the TOC field from the document.
                field.Remove();
            }
        }

        // Save the modified document as DOCM. The format is inferred from the file extension.
        string outputPath = @"C:\Output\DocumentWithoutToc.docm";
        doc.Save(outputPath);
    }
}
