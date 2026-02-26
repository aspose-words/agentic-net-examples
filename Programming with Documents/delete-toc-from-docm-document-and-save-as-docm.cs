using System;
using Aspose.Words;
using Aspose.Words.Fields;

class DeleteTocFromDocm
{
    static void Main()
    {
        // Path to the source DOCM file that contains a Table of Contents.
        string sourcePath = @"C:\Docs\SourceDocument.docm";

        // Path where the modified DOCM file will be saved.
        string destinationPath = @"C:\Docs\SourceDocument_NoToc.docm";

        // Load the existing DOCM document.
        Document doc = new Document(sourcePath);

        // Iterate through all fields in the document.
        // Remove each field that is a Table of Contents (TOC) field.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
            {
                // Cast to FieldToc to access TOC‑specific members if needed.
                FieldToc toc = (FieldToc)field;
                // Remove the TOC field from the document.
                toc.Remove();
            }
        }

        // Save the modified document back as a DOCM file.
        doc.Save(destinationPath);
    }
}
