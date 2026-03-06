using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Path to the DOTM template.
        string templatePath = "Template.dotm";

        // Load the document with conversion of EquationXML shapes to OfficeMath objects.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ConvertShapeToOfficeMath = true;
        Document doc = new Document(templatePath, loadOptions);

        // Retrieve all OfficeMath nodes in the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Prepare HTML save options to export OfficeMath as MathML (XML).
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            OfficeMathOutputMode = HtmlOfficeMathOutputMode.MathML
        };

        // Output the MathML XML of each equation to the console.
        foreach (OfficeMath officeMath in mathNodes)
        {
            string mathXml = officeMath.ToString(htmlOptions);
            Console.WriteLine(mathXml);
        }
    }
}
