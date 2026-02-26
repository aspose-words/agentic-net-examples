using System;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Path to the source DOTX template.
        string inputPath = @"C:\Docs\Template.dotx";

        // Path where the modified document will be saved.
        string outputPath = @"C:\Docs\Result.docx";

        // Load the DOTX document.
        Document doc = new Document(inputPath);

        // Retrieve all OfficeMath nodes in the document (including those inside paragraphs).
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath object and toggle its justification.
        foreach (Node node in mathNodes)
        {
            OfficeMath officeMath = (OfficeMath)node;

            // The Justification property can only be set when the display type is Display.
            officeMath.DisplayType = OfficeMathDisplayType.Display;

            // Simple toggle: Left <-> Center.
            if (officeMath.Justification == OfficeMathJustification.Left)
                officeMath.Justification = OfficeMathJustification.Center;
            else
                officeMath.Justification = OfficeMathJustification.Left;
        }

        // Save the modified document.
        doc.Save(outputPath);
    }
}
