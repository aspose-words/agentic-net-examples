using System;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load the document that contains the OfficeMath node.
        Document doc = new Document("Input.docx");

        // Retrieve the first OfficeMath node in the document (change the index if needed).
        OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
        if (officeMath != null)
        {
            // Keep a reference to the paragraph that hosts the OfficeMath node.
            Paragraph parentParagraph = officeMath.ParentParagraph;

            // Remove the OfficeMath node from its parent paragraph.
            officeMath.Remove();

            // If the paragraph still exists, adjust its spacing to compensate for the removed equation.
            if (parentParagraph != null)
            {
                // Example: set 12 points of space before and after the paragraph.
                parentParagraph.ParagraphFormat.SpaceBefore = 12;
                parentParagraph.ParagraphFormat.SpaceAfter = 12;
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
