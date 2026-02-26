using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Load an existing DOC document from disk.
        // The Document constructor is the provided lifecycle rule for loading.
        Document doc = new Document("Input.doc");

        // Define the style we want to target.
        // You can use either the style name or the locale‑independent identifier.
        const string targetStyleName = "Emphasis";                 // example style name
        const StyleIdentifier targetStyleId = StyleIdentifier.Emphasis; // same style by identifier

        // Iterate through all Run nodes in the document.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            // Check if the run uses the desired style (by name or identifier).
            bool matchesByName = string.Equals(run.Font.StyleName, targetStyleName, StringComparison.OrdinalIgnoreCase);
            bool matchesById   = run.Font.StyleIdentifier == targetStyleId;

            if (matchesByName || matchesById)
            {
                // Change the font color of the matching run.
                run.Font.Color = Color.Red;
            }
        }

        // Save the modified document.
        // The Document.Save method is the provided lifecycle rule for saving.
        doc.Save("Output.docx");
    }
}
