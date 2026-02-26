using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Path to the Word document that contains equations.
        string filePath = "MyDir\\MathDocument.docx";

        // Load options: convert shapes that contain EquationXML to OfficeMath objects.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ConvertShapeToOfficeMath = true;

        // Load the document with the specified options.
        Document doc = new Document(filePath, loadOptions);

        // Retrieve all OfficeMath nodes (equations) in the document.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node and output its plain‑text representation.
        for (int i = 0; i < officeMathNodes.Count; i++)
        {
            OfficeMath officeMath = (OfficeMath)officeMathNodes[i];
            string equationText = officeMath.GetText(); // Gets the text of the equation.
            Console.WriteLine($"Equation {i + 1}: {equationText.Trim()}");
        }
    }
}
