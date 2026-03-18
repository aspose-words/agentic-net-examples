using System;
using Aspose.Words;
using Aspose.Words.Math;

class UpdateOfficeMathJustification
{
    static void Main()
    {
        // Path to the template document that contains OfficeMath equations.
        const string templatePath = @"C:\Docs\Template.docx";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Retrieve all OfficeMath nodes in the document (including those inside other nodes).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node and set its justification to right.
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            // Ensure the equation is displayed on its own line; justification works only for Display type.
            officeMath.DisplayType = OfficeMathDisplayType.Display;

            // Set the justification to right alignment.
            officeMath.Justification = OfficeMathJustification.Right;
        }

        // Save the modified document.
        const string outputPath = @"C:\Docs\Template_JustifiedRight.docx";
        doc.Save(outputPath);
    }
}
