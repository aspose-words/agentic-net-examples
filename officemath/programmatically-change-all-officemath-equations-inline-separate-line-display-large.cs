using System;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Retrieve all OfficeMath nodes in the document (including nested ones).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node.
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            // We only change top‑level equations (MathObjectType.OMathPara) that are currently inline.
            if (officeMath.MathObjectType == MathObjectType.OMathPara &&
                officeMath.DisplayType == OfficeMathDisplayType.Inline)
            {
                // Switch the equation to display mode (its own line) and left‑justify it.
                officeMath.DisplayType = OfficeMathDisplayType.Display;
                officeMath.Justification = OfficeMathJustification.Left;
            }
        }

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
