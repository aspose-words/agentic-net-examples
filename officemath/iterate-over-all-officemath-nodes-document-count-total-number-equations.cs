using System;
using Aspose.Words;
using Aspose.Words.Math;

class CountOfficeMathEquations
{
    static void Main()
    {
        // Load the document. Replace the path with the actual file location.
        Document doc = new Document("Input.docx");

        // Get a collection of all OfficeMath nodes in the document (including nested ones).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Count the total number of equations (OfficeMath objects).
        int equationCount = officeMathNodes.Count;

        // Output the result.
        Console.WriteLine($"Total number of OfficeMath equations: {equationCount}");

        // Optionally, save the document (no changes made, just demonstrating the save lifecycle).
        doc.Save("Output.docx");
    }
}
