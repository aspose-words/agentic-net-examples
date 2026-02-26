using System;
using Aspose.Words;
using Aspose.Words.Notes;
using Aspose.Words.Saving;

class DeleteFootnoteExample
{
    static void Main()
    {
        // Path to the source DOCM file.
        string inputPath = @"C:\Docs\SourceDocument.docm";

        // Path where the resulting DOT template will be saved.
        string outputPath = @"C:\Docs\ResultTemplate.dot";

        // Load the DOCM document.
        Document doc = new Document(inputPath);

        // OPTIONAL: Update fields so that reference marks are current.
        doc.UpdateFields();
        doc.UpdateActualReferenceMarks();

        // Identify the footnote to delete.
        // For example, delete the footnote with index 0 (the first footnote in the document).
        // Adjust the index as needed or locate by custom ReferenceMark.
        Footnote footnote = doc.GetChild(NodeType.Footnote, 0, true) as Footnote;

        if (footnote != null)
        {
            // Remove the footnote from its parent.
            footnote.Remove();
        }

        // Save the modified document as a DOT template.
        doc.Save(outputPath, SaveFormat.Dot);
    }
}
