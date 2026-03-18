using System;
using Aspose.Words;
using Aspose.Words.Math;

class CloneOfficeMathExample
{
    static void Main()
    {
        // Load the source document that contains an OfficeMath object.
        Document doc = new Document("OfficeMathSource.docx");

        // Locate the first OfficeMath node in the document.
        OfficeMath originalMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
        if (originalMath == null)
        {
            Console.WriteLine("No OfficeMath object found in the document.");
            return;
        }

        // Clone the OfficeMath node together with all its child nodes.
        // The Clone method returns a Node; cast it back to OfficeMath.
        OfficeMath clonedMath = (OfficeMath)originalMath.Clone(true);

        // Choose a location where the cloned OfficeMath will be inserted.
        // For this example, insert it after the paragraph that follows the original math.
        Paragraph referenceParagraph = originalMath.ParentParagraph;
        Paragraph targetParagraph = referenceParagraph.NextSibling as Paragraph;
        if (targetParagraph == null)
        {
            // If there is no following paragraph, create a new one at the end of the body.
            targetParagraph = (Paragraph)doc.FirstSection.Body.AppendChild(new Paragraph(doc));
        }

        // Insert the cloned OfficeMath after the target paragraph.
        // Since the cloned node belongs to the same document, no import is required.
        targetParagraph.ParentNode.InsertAfter(clonedMath, targetParagraph);

        // Optionally, adjust display settings of the cloned math.
        clonedMath.DisplayType = OfficeMathDisplayType.Display;
        clonedMath.Justification = OfficeMathJustification.Left;

        // Save the modified document.
        doc.Save("OfficeMathCloned.docx");
    }
}
