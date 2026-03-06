using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Replacing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("Input.doc");

        // Define the style we are looking for (by name or by identifier).
        const string targetStyleName = "Emphasis";                     // Example style name.
        StyleIdentifier targetStyleId = StyleIdentifier.Emphasis;    // Example built‑in style identifier.

        // Iterate through all Run nodes in the document.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            // Check if the run uses the target style (either by name or by identifier).
            bool hasTargetStyle = run.Font.StyleName == targetStyleName ||
                                  run.Font.StyleIdentifier == targetStyleId;

            if (hasTargetStyle)
            {
                // Change the font color of the run.
                run.Font.Color = Color.Red;
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
