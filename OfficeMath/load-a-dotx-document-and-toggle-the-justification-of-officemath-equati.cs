using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;

class ToggleOfficeMathJustification
{
    static void Main()
    {
        // Path to the folder that contains the DOTX template.
        string dataDir = @"C:\Docs\Templates";
        string inputPath = Path.Combine(dataDir, "Template.dotx");

        // Load the DOTX document.
        Document doc = new Document(inputPath);

        // Retrieve all OfficeMath nodes in the document (including those in headers/footers).
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath object and toggle its justification.
        foreach (OfficeMath officeMath in mathNodes)
        {
            // Ensure the equation is displayed on its own line before changing justification.
            officeMath.DisplayType = OfficeMathDisplayType.Display;

            // Toggle between Left and Center justification.
            // If the current justification is Left, change to Center; otherwise set to Left.
            if (officeMath.Justification == OfficeMathJustification.Left)
                officeMath.Justification = OfficeMathJustification.Center;
            else
                officeMath.Justification = OfficeMathJustification.Left;
        }

        // Save the modified document.
        string outputPath = Path.Combine(dataDir, "Template_Toggled.docx");
        doc.Save(outputPath);
    }
}
