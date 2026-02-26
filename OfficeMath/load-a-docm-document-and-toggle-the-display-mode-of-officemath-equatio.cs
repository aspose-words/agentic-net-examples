using System;
using Aspose.Words;
using Aspose.Words.Math;

class ToggleOfficeMathDisplay
{
    static void Main()
    {
        // Load the DOCM document.
        Document doc = new Document("Input.docm");

        // Retrieve all OfficeMath nodes in the document (including nested ones).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Toggle the display type for each top‑level OfficeMath node.
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            // Only top‑level equations have a mutable DisplayType.
            // Nested OfficeMath objects are always inline and cannot be changed.
            if (officeMath.MathObjectType == MathObjectType.OMathPara)
            {
                // Switch between Inline and Display.
                officeMath.DisplayType = officeMath.DisplayType == OfficeMathDisplayType.Inline
                    ? OfficeMathDisplayType.Display
                    : OfficeMathDisplayType.Inline;
            }
        }

        // Save the modified document.
        doc.Save("Output.docm");
    }
}
