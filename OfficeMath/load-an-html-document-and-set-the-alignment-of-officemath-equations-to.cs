using System;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source HTML file.
        string htmlPath = "input.html";

        // Load the HTML document. Enable conversion of EquationXML shapes to OfficeMath objects.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ConvertShapeToOfficeMath = true;
        Document doc = new Document(htmlPath, loadOptions);

        // Iterate through all OfficeMath nodes in the document.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            // The justification can be set only when the display type is Display.
            officeMath.DisplayType = OfficeMathDisplayType.Display;

            // Center each equation.
            officeMath.Justification = OfficeMathJustification.Center;
        }

        // Save the modified document.
        doc.Save("output.docx");
    }
}
