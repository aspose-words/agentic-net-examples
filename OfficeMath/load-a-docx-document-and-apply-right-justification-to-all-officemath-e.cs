using System;
using Aspose.Words;
using Aspose.Words.Math;

class ApplyRightJustificationToOfficeMath
{
    static void Main()
    {
        // Load the existing DOCX document.
        string inputPath = "input.docx";
        Document doc = new Document(inputPath);

        // Retrieve all OfficeMath nodes in the document (including nested ones).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Apply right justification to each OfficeMath equation.
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            // Ensure the equation is displayed on its own line before setting justification.
            officeMath.DisplayType = OfficeMathDisplayType.Display;

            // Set the justification to right.
            officeMath.Justification = OfficeMathJustification.Right;
        }

        // Save the modified document.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
