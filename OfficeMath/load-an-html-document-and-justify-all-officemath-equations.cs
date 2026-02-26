using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load the HTML document using HtmlLoadOptions.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        Document doc = new Document("Input.html", loadOptions);

        // Iterate through all OfficeMath nodes in the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath officeMath in mathNodes)
        {
            // Ensure the equation is displayed on its own line before setting justification.
            officeMath.DisplayType = OfficeMathDisplayType.Display;

            // Apply the desired justification to the equation.
            officeMath.Justification = OfficeMathJustification.CenterGroup;
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
