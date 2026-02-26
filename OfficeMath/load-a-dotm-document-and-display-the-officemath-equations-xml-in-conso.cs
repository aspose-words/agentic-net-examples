using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Path to the DOTM template that contains OfficeMath equations.
        string templatePath = @"C:\Docs\Template.dotm";

        // Load the document with conversion of shapes that have EquationXML to OfficeMath objects.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ConvertShapeToOfficeMath = true;
        Document doc = new Document(templatePath, loadOptions);

        // Prepare HTML save options that export OfficeMath as MathML (XML).
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.OfficeMathOutputMode = HtmlOfficeMathOutputMode.MathML;

        // Retrieve all OfficeMath nodes in the document.
        int mathCount = doc.GetChildNodes(NodeType.OfficeMath, true).Count;

        for (int i = 0; i < mathCount; i++)
        {
            // Cast the node to OfficeMath.
            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, i, true);

            // Export the OfficeMath node to a string containing its MathML representation.
            string mathMl = officeMath.ToString(htmlOptions);

            // Output the XML to the console.
            Console.WriteLine($"OfficeMath #{i + 1} XML:");
            Console.WriteLine(mathMl);
            Console.WriteLine();
        }
    }
}
