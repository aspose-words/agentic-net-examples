using System;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load the Markdown document.
        // Aspose.Words can directly load .md files.
        Document doc = new Document("input.md");

        // Retrieve all OfficeMath nodes in the document (including nested ones).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node and output its MathObjectType.
        for (int i = 0; i < officeMathNodes.Count; i++)
        {
            OfficeMath officeMath = (OfficeMath)officeMathNodes[i];
            Console.WriteLine($"OfficeMath #{i}: {officeMath.MathObjectType}");
        }
    }
}
