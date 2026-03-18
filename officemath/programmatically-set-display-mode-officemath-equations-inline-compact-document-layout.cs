using System;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load the source document.
        // Replace "Input.docx" with the actual path to your document.
        Document doc = new Document("Input.docx");

        // Retrieve all OfficeMath nodes in the document (including nested ones).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node.
        foreach (Node node in officeMathNodes)
        {
            OfficeMath officeMath = (OfficeMath)node;

            // Only top‑level OfficeMath objects (MathObjectType.OMathPara) allow changing the display type.
            if (officeMath.MathObjectType == MathObjectType.OMathPara)
            {
                // Set the equation to be displayed inline for a more compact layout.
                officeMath.DisplayType = OfficeMathDisplayType.Inline;

                // When the display type is Inline, the justification must also be Inline.
                officeMath.Justification = OfficeMathJustification.Inline;
            }
        }

        // Save the modified document.
        // Replace "Output.docx" with the desired output path.
        doc.Save("Output.docx");
    }
}
