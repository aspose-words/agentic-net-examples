using System;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Loading;

class ListOfficeMath
{
    static void Main()
    {
        // Path to the DOCX file.
        string filePath = "input.docx";

        // Load the document with shapes converted to OfficeMath objects.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ConvertShapeToOfficeMath = true;
        Document doc = new Document(filePath, loadOptions);

        // Retrieve all OfficeMath nodes in the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        Console.WriteLine($"Found {mathNodes.Count} OfficeMath equation(s):");

        // List each equation's type and its plain text representation.
        foreach (OfficeMath officeMath in mathNodes)
        {
            string text = officeMath.GetText().Trim();
            Console.WriteLine($"- Type: {officeMath.MathObjectType}, Text: \"{text}\"");
        }
    }
}
