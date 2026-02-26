using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Drawing;

class RetrieveOfficeMathFontSizes
{
    static void Main()
    {
        // Path to the DOCX file that contains OfficeMath equations.
        string docPath = @"C:\Docs\SampleWithMath.docx";

        // Load the document from the file system.
        Document doc = new Document(docPath);

        // Get all OfficeMath nodes in the document (including those inside tables, headers, etc.).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node and obtain the font size.
        int equationIndex = 0;
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            // Retrieve the first Run node inside the OfficeMath element.
            // OfficeMath can contain multiple Run nodes; typically they share the same font size.
            Run firstRun = officeMath.GetChildNodes(NodeType.Run, true)
                                      .Cast<Run>()
                                      .FirstOrDefault();

            // If a Run is found, read its Font.Size; otherwise default to 0.
            double fontSize = firstRun?.Font?.Size ?? 0;

            Console.WriteLine($"Equation {equationIndex}: Font size = {fontSize} pt");
            equationIndex++;
        }
    }
}
