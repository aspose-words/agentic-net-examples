using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Path to the source Word document.
        string inputFile = @"C:\Docs\SourceDocument.docx";

        // Path to the text file that will contain the extracted equations.
        string outputFile = @"C:\Docs\ExtractedEquations.txt";

        // Load the document using the standard constructor (lifecycle rule).
        Document doc = new Document(inputFile);

        // Collect all OfficeMath nodes in the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Build a string with each equation on a separate line.
        StringBuilder sb = new StringBuilder();

        foreach (OfficeMath math in mathNodes)
        {
            // Get the plain‑text representation of the equation.
            // OfficeMath.GetText() returns the text of the node and its children.
            string equation = math.GetText().Trim();

            if (!string.IsNullOrEmpty(equation))
            {
                sb.AppendLine(equation);
            }
        }

        // Write the collected equations to the output text file.
        File.WriteAllText(outputFile, sb.ToString());
    }
}
