using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load the HTML document.
        var loadOptions = new HtmlLoadOptions();
        Document doc = new Document("input.html", loadOptions);

        // Retrieve all OfficeMath nodes in the document.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node and set its display type to Display.
        // (MathObjectType is read‑only; setting DisplayType changes how the equation is rendered,
        // which is the closest operation to “integral” formatting in this context.)
        foreach (OfficeMath om in officeMathNodes)
        {
            om.DisplayType = OfficeMathDisplayType.Display;
        }

        // Save the modified document.
        doc.Save("output.docx");
    }
}
