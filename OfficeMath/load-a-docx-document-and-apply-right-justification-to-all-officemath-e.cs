using Aspose.Words;
using Aspose.Words.Math;
using System;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Retrieve all OfficeMath objects in the document.
        NodeCollection officeMaths = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Apply right justification to each equation.
        foreach (OfficeMath om in officeMaths)
        {
            // Set display type to Display before changing justification.
            om.DisplayType = OfficeMathDisplayType.Display;
            om.Justification = OfficeMathJustification.Right;
        }

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
