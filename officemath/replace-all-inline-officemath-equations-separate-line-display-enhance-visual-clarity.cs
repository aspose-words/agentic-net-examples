using System;
using Aspose.Words;
using Aspose.Words.Math;

class InlineOfficeMathToDisplay
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Get all OfficeMath nodes in the document (including those inside other nodes).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node and change its display type if it is inline.
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            // Only top‑level OfficeMath objects have a mutable DisplayType.
            if (officeMath.DisplayType == OfficeMathDisplayType.Inline)
            {
                officeMath.DisplayType = OfficeMathDisplayType.Display;
                // Optional: set left justification for the displayed equation.
                officeMath.Justification = OfficeMathJustification.Left;
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
