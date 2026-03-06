using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Notes;

class DeleteFootnoteExample
{
    static void Main()
    {
        // Path to the source DOCM file.
        string inputPath = @"C:\Docs\SourceDocument.docm";

        // Path where the modified DOCM will be saved.
        string outputPath = @"C:\Docs\ModifiedDocument.docm";

        // Load the DOCM document.
        Document doc = new Document(inputPath);

        // Index of the footnote to delete (0‑based). 
        // For example, to delete the first footnote use index = 0.
        int footnoteIndex = 0;

        // Retrieve the footnote node. The GetChild method searches the whole document.
        Footnote footnote = (Footnote)doc.GetChild(NodeType.Footnote, footnoteIndex, true);

        if (footnote != null)
        {
            // Remove the footnote from its parent paragraph.
            footnote.Remove();

            // After removal, update the actual reference marks so that remaining footnotes are renumbered correctly.
            doc.UpdateFields();
            doc.UpdateActualReferenceMarks();
        }

        // Save the modified document as DOCM, preserving the macro-enabled format.
        doc.Save(outputPath, SaveFormat.Docm);
    }
}
