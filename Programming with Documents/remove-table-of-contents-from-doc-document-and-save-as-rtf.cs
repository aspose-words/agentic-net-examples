using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class RemoveTocAndSaveAsRtf
{
    static void Main()
    {
        // Path to the source DOC/DOCX file.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the resulting RTF file will be saved.
        string outputPath = @"C:\Docs\ResultDocument.rtf";

        // Load the document from disk.
        Document doc = new Document(inputPath);

        // Iterate through all field start nodes in the document.
        // If a field start belongs to a Table of Contents (TOC) field, remove the entire field.
        foreach (FieldStart fieldStart in doc.GetChildNodes(NodeType.FieldStart, true))
        {
            if (fieldStart.FieldType == FieldType.FieldTOC)
            {
                // Retrieve the full field (start, separator, end) and remove it.
                Field tocField = fieldStart.GetField();
                tocField?.Remove();
            }
        }

        // Create RTF save options (optional customizations can be set here).
        RtfSaveOptions rtfOptions = new RtfSaveOptions
        {
            // Example: reduce file size when RTL text is not required.
            ExportCompactSize = true
        };

        // Save the modified document as RTF using the specified options.
        doc.Save(outputPath, rtfOptions);
    }
}
