using System;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load the DOTM document.
        Document doc = new Document("Template.dotm");

        // Retrieve all OfficeMath nodes in the document (deep search).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node.
        for (int i = 0; i < officeMathNodes.Count; i++)
        {
            OfficeMath officeMath = (OfficeMath)officeMathNodes[i];

            // Get the paragraph that contains this OfficeMath node.
            Paragraph parentParagraph = officeMath.ParentParagraph;

            // Determine the paragraph's index within the document.
            int paragraphIndex = -1;
            if (parentParagraph != null)
            {
                NodeCollection allParagraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                paragraphIndex = allParagraphs.IndexOf(parentParagraph);
            }

            // Output the OfficeMath information.
            Console.WriteLine($"OfficeMath #{i}: Type = {officeMath.MathObjectType}, ParagraphIndex = {paragraphIndex}");
        }
    }
}
