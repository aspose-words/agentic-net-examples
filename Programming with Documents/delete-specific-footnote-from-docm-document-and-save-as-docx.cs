using System;
using Aspose.Words;
using Aspose.Words.Notes;

class DeleteFootnoteExample
{
    static void Main()
    {
        // Path to the source DOCM file.
        const string inputPath = @"C:\Docs\SourceDocument.docm";

        // Path where the resulting DOCX will be saved.
        const string outputPath = @"C:\Docs\ResultDocument.docx";

        // Zero‑based index of the footnote to delete.
        // Change this value to target a different footnote.
        const int footnoteIndexToDelete = 0;

        // Load the DOCM document.
        Document doc = new Document(inputPath);

        // Retrieve all footnotes in the document (including those in headers/footers).
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);

        // Ensure the requested index exists.
        if (footnoteIndexToDelete >= 0 && footnoteIndexToDelete < footnotes.Count)
        {
            // Cast the node to a Footnote and remove it from its parent.
            Footnote footnote = (Footnote)footnotes[footnoteIndexToDelete];
            footnote.Remove();
        }
        else
        {
            Console.WriteLine("Footnote index out of range.");
        }

        // Save the modified document as DOCX.
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
