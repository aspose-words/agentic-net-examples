using Aspose.Words;
using Aspose.Words.Math;
using System;

class ToggleOfficeMathDisplay
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Retrieve all OfficeMath nodes in the document.
        NodeCollection officeMaths = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node.
        foreach (OfficeMath om in officeMaths)
        {
            // Only top‑level OfficeMath objects (MathObjectType.OMathPara) support display type changes.
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                // Toggle between Display (on its own line) and Inline (within text).
                if (om.DisplayType == OfficeMathDisplayType.Display)
                {
                    om.DisplayType = OfficeMathDisplayType.Inline;
                    // Inline equations do not use justification.
                }
                else
                {
                    om.DisplayType = OfficeMathDisplayType.Display;
                    // When displayed on its own line, set left justification.
                    om.Justification = OfficeMathJustification.Left;
                }
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
