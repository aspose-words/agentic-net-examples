using System;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Load the HTML document. Aspose.Words automatically detects the format.
        Document doc = new Document("input.html");

        // Iterate through all OfficeMath objects in the document.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            // Ensure the equation is displayed on its own line before setting justification.
            officeMath.DisplayType = OfficeMathDisplayType.Display;

            // Center the equation within the page.
            officeMath.Justification = OfficeMathJustification.Center;
        }

        // Save the modified document.
        doc.Save("output.docx");
    }
}
