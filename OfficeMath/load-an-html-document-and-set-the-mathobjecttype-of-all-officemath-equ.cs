using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load the HTML document. HtmlLoadOptions can be customized if needed.
        var loadOptions = new HtmlLoadOptions();
        Document doc = new Document("input.html", loadOptions);

        // Iterate through all OfficeMath nodes in the document.
        // The MathObjectType property is read‑only, so we cannot change it directly.
        // Instead, we set the display format to Inline, which is the typical
        // representation for an integral (inline) equation.
        foreach (OfficeMath officeMath in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            // Set the equation to be displayed inline.
            officeMath.DisplayType = OfficeMathDisplayType.Inline;
        }

        // Save the modified document.
        doc.Save("output.html");
    }
}
