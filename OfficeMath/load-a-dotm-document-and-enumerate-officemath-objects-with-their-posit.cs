using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Tables;

class EnumerateOfficeMath
{
    static void Main()
    {
        // Path to the folder that contains the DOTM template.
        string dataDir = @"C:\Docs\";

        // Load the DOTM document.
        Document doc = new Document(Path.Combine(dataDir, "Template.dotm"));

        // Get all OfficeMath nodes in the document (including those inside other OfficeMath nodes).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Enumerate each OfficeMath node and output its type and its containing paragraph index.
        for (int i = 0; i < officeMathNodes.Count; i++)
        {
            OfficeMath officeMath = (OfficeMath)officeMathNodes[i];

            // Determine the index of the paragraph that contains this OfficeMath node.
            Paragraph parentParagraph = officeMath.ParentParagraph;
            int paragraphIndex = -1;
            if (parentParagraph != null)
            {
                NodeCollection allParagraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                paragraphIndex = allParagraphs.IndexOf(parentParagraph);
            }

            Console.WriteLine($"OfficeMath #{i}: Type = {officeMath.MathObjectType}, ParagraphIndex = {paragraphIndex}");
        }

        // Optionally save the document after enumeration (preserves original format).
        doc.Save(Path.Combine(dataDir, "Template_out.dotm"));
    }
}
