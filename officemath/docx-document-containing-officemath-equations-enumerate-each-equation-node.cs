using System;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load the DOCX document that contains OfficeMath equations.
        // The Document constructor is the provided load rule.
        Document doc = new Document("InputWithOfficeMath.docx");

        // Retrieve all OfficeMath nodes in the document (including those in child nodes).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Enumerate each OfficeMath node and output its index and type.
        for (int i = 0; i < officeMathNodes.Count; i++)
        {
            OfficeMath officeMath = (OfficeMath)officeMathNodes[i];
            Console.WriteLine($"Equation {i + 1}: MathObjectType = {officeMath.MathObjectType}");
        }

        // (Optional) Save the document if any modifications were made.
        // The Document.Save method is the provided save rule.
        // doc.Save("Output.docx");
    }
}
