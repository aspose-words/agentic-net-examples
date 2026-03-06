using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields; // Added namespace for Field and FieldType

class Program
{
    static void Main()
    {
        // Paths to the source DOC document and the destination TXT file.
        string inputPath = "Input.docx";
        string outputPath = "Output.txt";

        // Load the Word document from disk.
        Document doc = new Document(inputPath);

        // Remove every Table of Contents (TOC) field from the document.
        // Iterate backwards because removing a field modifies the collection.
        for (int i = doc.Range.Fields.Count - 1; i >= 0; i--)
        {
            Field field = doc.Range.Fields[i];
            if (field.Type == FieldType.FieldTOC)
                field.Remove();
        }

        // Configure plain‑text save options (optional: omit headers/footers).
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.None
        };

        // Save the modified document as a .txt file.
        doc.Save(outputPath, txtOptions);
    }
}
