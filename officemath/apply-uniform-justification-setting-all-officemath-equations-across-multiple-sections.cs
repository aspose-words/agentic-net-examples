using System;
using Aspose.Words;
using Aspose.Words.Math;

class ApplyOfficeMathJustification
{
    static void Main()
    {
        // Path to the source document.
        const string inputPath = @"C:\Docs\InputDocument.docx";

        // Path where the modified document will be saved.
        const string outputPath = @"C:\Docs\OutputDocument.docx";

        // Load the document (lifecycle rule: load).
        Document doc = new Document(inputPath);

        // Retrieve all OfficeMath nodes in the document, including those in all sections.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Desired uniform justification for every equation.
        OfficeMathJustification targetJustification = OfficeMathJustification.CenterGroup;

        // Apply the justification to each OfficeMath node.
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            // Justification cannot be set while the equation is Inline.
            // Ensure the equation is displayed on its own line first.
            if (officeMath.DisplayType == OfficeMathDisplayType.Inline)
                officeMath.DisplayType = OfficeMathDisplayType.Display;

            // Set the uniform justification.
            officeMath.Justification = targetJustification;
        }

        // Save the modified document (lifecycle rule: save).
        doc.Save(outputPath);
    }
}
