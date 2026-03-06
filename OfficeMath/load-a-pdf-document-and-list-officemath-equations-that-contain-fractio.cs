using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main(string[] args)
    {
        // Path to the PDF file to be loaded.
        string pdfPath = "input.pdf";

        // Load the PDF document. Aspose.Words automatically detects the format.
        Document doc = new Document(pdfPath);

        // Retrieve all OfficeMath nodes in the document (including those inside other OfficeMath nodes).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node.
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            // Determine if the current OfficeMath node (or any of its descendants) contains a Fraction object.
            bool containsFraction = officeMath
                .GetChildNodes(NodeType.OfficeMath, true)               // Get all descendant OfficeMath nodes.
                .Cast<OfficeMath>()                                    // Cast to OfficeMath for LINQ.
                .Any(child => child.MathObjectType == MathObjectType.Fraction);

            // If a fraction is present, output the equation's plain text.
            if (containsFraction)
            {
                Console.WriteLine("Equation containing a fraction:");
                Console.WriteLine(officeMath.GetText().Trim());
                Console.WriteLine(); // Blank line for readability.
            }
        }
    }
}
