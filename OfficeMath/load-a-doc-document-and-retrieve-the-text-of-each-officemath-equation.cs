using System;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the input Word document.
        const string inputPath = "input.docx";

        // Load the document with the option to convert EquationXML shapes to OfficeMath objects.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ConvertShapeToOfficeMath = true; // Ensure all equations are loaded as OfficeMath.
        Document doc = new Document(inputPath, loadOptions);

        // Retrieve all OfficeMath nodes in the document (including those inside other nodes).
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node and output its plain‑text representation.
        for (int i = 0; i < mathNodes.Count; i++)
        {
            OfficeMath officeMath = (OfficeMath)mathNodes[i];
            string equationText = officeMath.GetText().Trim(); // Get the equation text.
            Console.WriteLine($"Equation {i + 1}: {equationText}");
        }
    }
}
