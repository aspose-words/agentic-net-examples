using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Math;

namespace OfficeMathInlineFinder
{
    class Program
    {
        static void Main()
        {
            // Path to the WORDML (or any supported) document.
            string inputPath = "input.docx";

            // Load the document using the standard constructor.
            Document doc = new Document(inputPath);

            // Retrieve all OfficeMath nodes in the document (including nested ones).
            NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

            // List to hold inline OfficeMath equations.
            List<OfficeMath> inlineEquations = new List<OfficeMath>();

            // Iterate through each OfficeMath node and check its DisplayType.
            foreach (OfficeMath om in officeMathNodes)
            {
                // Only top‑level OfficeMath objects have a meaningful DisplayType.
                // Nested OfficeMath objects are always inline, but we are interested in
                // equations whose DisplayType is explicitly set to Inline.
                if (om.DisplayType == OfficeMathDisplayType.Inline)
                {
                    inlineEquations.Add(om);
                }
            }

            // Output the results.
            Console.WriteLine($"Total OfficeMath objects found: {officeMathNodes.Count}");
            Console.WriteLine($"Inline OfficeMath equations found: {inlineEquations.Count}");

            // Optionally, display the plain text of each inline equation.
            for (int i = 0; i < inlineEquations.Count; i++)
            {
                Console.WriteLine($"Inline Equation {i + 1}: {inlineEquations[i].GetText().Trim()}");
            }
        }
    }
}
