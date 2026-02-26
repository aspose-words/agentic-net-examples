using System;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load the WORDML document.
        // Replace the path with the actual location of your document.
        string filePath = "Input.docx";
        Document doc = new Document(filePath);

        // Retrieve all OfficeMath nodes in the document.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Identify and output inline equations.
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            if (officeMath.DisplayType == OfficeMathDisplayType.Inline)
            {
                // Example output: type of math object and its plain text.
                Console.WriteLine($"Inline equation - Type: {officeMath.MathObjectType}, Text: \"{officeMath.GetText().Trim()}\"");
            }
        }
    }
}
