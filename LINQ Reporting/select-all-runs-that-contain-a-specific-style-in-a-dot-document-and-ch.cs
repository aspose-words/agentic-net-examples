using System;
using System.Drawing;
using Aspose.Words;

class ChangeRunStyleColor
{
    static void Main()
    {
        // Load an existing DOC/DOCX document.
        Document doc = new Document("Input.docx");

        // Define the style name to look for.
        string targetStyleName = "Emphasis";

        // Define the new font color.
        Color newColor = Color.Blue;

        // Iterate through all Run nodes in the document.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            // Check if the run uses the specified style name.
            if (run.Font.StyleName == targetStyleName)
            {
                // Change the font color of the run.
                run.Font.Color = newColor;
            }

            // Alternatively, you can check by built‑in style identifier:
            // if (run.Font.StyleIdentifier == StyleIdentifier.Emphasis)
            // {
            //     run.Font.Color = newColor;
            // }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
