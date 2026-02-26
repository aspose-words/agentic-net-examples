using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

class OfficeMathExample
{
    static void Main()
    {
        // Path to the source DOCX file that contains EquationXML shapes.
        string inputPath = @"C:\Docs\MathShapes.docx";

        // Configure load options to convert EquationXML shapes to OfficeMath objects.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ConvertShapeToOfficeMath = true;

        // Load the document with the specified options.
        Document doc = new Document(inputPath, loadOptions);

        // Example: count the OfficeMath nodes in the document.
        int officeMathCount = doc.GetChildNodes(NodeType.OfficeMath, true).Count;
        Console.WriteLine($"Number of OfficeMath objects: {officeMathCount}");

        // (Optional) Save the document after conversion for verification.
        string outputPath = @"C:\Docs\MathShapes_Converted.docx";
        doc.Save(outputPath);
    }
}
