using System;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load the DOCX document (uses the Document(string) constructor rule).
        Document doc = new Document("OfficeMath.docx");

        // Retrieve all OfficeMath nodes in the document.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath equation.
        for (int i = 0; i < officeMathNodes.Count; i++)
        {
            OfficeMath officeMath = (OfficeMath)officeMathNodes[i];

            // Get all Run nodes that are descendants of the current OfficeMath node.
            NodeCollection runs = officeMath.GetChildNodes(NodeType.Run, true);

            Console.WriteLine($"OfficeMath #{i + 1} contains {runs.Count} run(s):");

            // Output the font size for each Run.
            for (int j = 0; j < runs.Count; j++)
            {
                Run run = (Run)runs[j];
                double fontSize = run.Font.Size; // Font size in points.
                string text = run.GetText().Trim();
                Console.WriteLine($"  Run {j + 1}: Text=\"{text}\", FontSize={fontSize} pt");
            }
        }
    }
}
