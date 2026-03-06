using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Math;   // Provides OfficeMath node type

class OfficeMathExample
{
    static void Main()
    {
        // Path to the input DOCX file that contains equation shapes.
        string inputPath = @"C:\Docs\MathShapes.docx";

        // Configure load options to convert shapes that contain EquationXML into OfficeMath objects.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ConvertShapeToOfficeMath = true;

        // Load the document with the specified options.
        Document doc = new Document(inputPath, loadOptions);

        // Example operation: count the OfficeMath objects in the document.
        int officeMathCount = doc.GetChildNodes(NodeType.OfficeMath, true).Count;
        Console.WriteLine($"Number of OfficeMath objects: {officeMathCount}");

        // Iterate through each OfficeMath node (optional processing can be added here).
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            // For demonstration, output the plain text representation of each equation.
            Console.WriteLine(om.GetText().Trim());
        }

        // Save the processed document (if any modifications were made).
        string outputPath = @"C:\Docs\MathShapes_Processed.docx";
        doc.Save(outputPath);
    }
}
