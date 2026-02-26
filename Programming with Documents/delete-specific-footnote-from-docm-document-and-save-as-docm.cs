using System;
using Aspose.Words;
using Aspose.Words.Notes;

class Program
{
    static void Main()
    {
        // Paths to the source and destination DOCM files.
        string inputPath = "input.docm";
        string outputPath = "output.docm";

        // Load the existing DOCM document.
        Document doc = new Document(inputPath);

        // Example 1: Remove a footnote by its reference mark.
        // (If you know the exact reference mark of the footnote you want to delete.)
        foreach (Footnote footnote in doc.GetChildNodes(NodeType.Footnote, true))
        {
            if (footnote.ReferenceMark == "RefMark") // replace "RefMark" with the actual mark.
            {
                footnote.Remove();
                break; // Assuming only one footnote with this mark.
            }
        }

        // Example 2: Remove a footnote by its index (zero‑based).
        // Uncomment the following lines if you prefer to delete, for example, the first footnote.
        //Footnote footnoteByIndex = (Footnote)doc.GetChild(NodeType.Footnote, 0, true);
        //if (footnoteByIndex != null)
        //    footnoteByIndex.Remove();

        // Save the modified document as DOCM, preserving macros.
        doc.Save(outputPath, SaveFormat.Docm);
    }
}
