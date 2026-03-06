using System;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Path to the source DOCM file.
        string inputPath = "Input.docm";

        // Load the document using the built‑in Document constructor (lifecycle rule).
        Document doc = new Document(inputPath);

        // Retrieve all OfficeMath nodes in the document (including nested ones).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node and toggle its display type.
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            // Only top‑level OfficeMath objects have a display type that can be changed.
            // Nested OfficeMath objects are always inline and should be left untouched.
            if (officeMath.MathObjectType == MathObjectType.OMathPara)
            {
                // Switch between Inline and Display.
                officeMath.DisplayType = officeMath.DisplayType == OfficeMathDisplayType.Inline
                    ? OfficeMathDisplayType.Display
                    : OfficeMathDisplayType.Inline;
            }
        }

        // Path for the modified DOCM file.
        string outputPath = "Output.docm";

        // Save the document using the built‑in Save method (lifecycle rule).
        doc.Save(outputPath);
    }
}
