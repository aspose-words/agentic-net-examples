using System;
using Aspose.Words;
using Aspose.Words.Notes;

class DeleteFootnoteExample
{
    static void Main()
    {
        // Path to the source DOCM file.
        string inputPath = "input.docm";

        // Path where the resulting DOT file will be saved.
        string outputPath = "output.dot";

        // Load the existing DOCM document.
        Document doc = new Document(inputPath);

        // Index of the footnote to delete (zero‑based).
        // Adjust this value to target the specific footnote you need to remove.
        int footnoteIndex = 0;

        // Retrieve the footnote node. The GetChild method searches the whole document
        // for nodes of the specified type. The third parameter 'true' enables deep search.
        Footnote footnote = doc.GetChild(NodeType.Footnote, footnoteIndex, true) as Footnote;

        // If the footnote exists, remove it from its parent.
        if (footnote != null)
        {
            footnote.Remove();
        }

        // Save the modified document as a Word template (DOT format).
        doc.Save(outputPath, SaveFormat.Dot);
    }
}
