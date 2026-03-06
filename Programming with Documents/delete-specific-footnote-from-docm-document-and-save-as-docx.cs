using System;
using Aspose.Words;
using Aspose.Words.Notes;

class Program
{
    static void Main()
    {
        // Path to the source DOCM file.
        string inputPath = @"C:\Docs\SourceDocument.docm";

        // Path where the resulting DOCX will be saved.
        string outputPath = @"C:\Docs\SourceDocument_NoFootnote.docx";

        // Load the DOCM document using the constructor that accepts a file name.
        Document doc = new Document(inputPath);

        // Zero‑based index of the footnote to delete.
        // Adjust this value to target the required footnote.
        int footnoteIndexToDelete = 1;

        // Retrieve all footnotes in the document (including those in headers/footers if any).
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);

        // Verify that the requested footnote exists.
        if (footnoteIndexToDelete >= 0 && footnoteIndexToDelete < footnotes.Count)
        {
            // Cast the node to a Footnote and remove it from its parent.
            Footnote footnote = (Footnote)footnotes[footnoteIndexToDelete];
            footnote.Remove();
        }

        // Save the modified document as DOCX.
        // The Save(string) overload automatically selects the format based on the file extension.
        doc.Save(outputPath);
    }
}
