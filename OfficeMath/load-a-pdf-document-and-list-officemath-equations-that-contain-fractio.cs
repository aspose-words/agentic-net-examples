using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load the PDF document (Aspose.Words automatically detects the format).
        Document doc = new Document("input.pdf");

        // Retrieve all OfficeMath nodes in the document.
        var officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node.
        foreach (OfficeMath om in officeMathNodes)
        {
            // Determine whether the equation contains a fraction element.
            bool containsFraction = om.GetChildNodes(NodeType.OfficeMath, true)
                                      .Cast<OfficeMath>()
                                      .Any(child => child.MathObjectType == MathObjectType.Fraction);

            if (containsFraction)
            {
                // Output the equation text that includes a fraction.
                Console.WriteLine("Equation containing fraction:");
                Console.WriteLine(om.GetText().Trim());
                Console.WriteLine(new string('-', 40));
            }
        }
    }
}
