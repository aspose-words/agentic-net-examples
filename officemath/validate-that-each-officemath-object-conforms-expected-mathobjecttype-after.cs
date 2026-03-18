using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Saving;

namespace OfficeMathValidationExample
{
    public class OfficeMathValidator
    {
        /// <summary>
        /// Loads a document, performs bulk updates on OfficeMath nodes,
        /// and validates that each node's MathObjectType matches the expected value.
        /// </summary>
        /// <param name="inputPath">Path to the source DOCX file.</param>
        /// <param name="outputPath">Path where the (potentially modified) document will be saved.</param>
        /// <param name="expectedTypes">
        /// Mapping of OfficeMath node index (zero‑based) to the expected MathObjectType.
        /// If an index is not present in the dictionary, the node is considered valid by default.
        /// </param>
        /// <returns>True if all nodes match the expected types; otherwise false.</returns>
        public static bool Validate(string inputPath, string outputPath, Dictionary<int, MathObjectType> expectedTypes)
        {
            // Load the document (lifecycle rule: load)
            Document doc = new Document(inputPath);

            // Perform bulk updates on OfficeMath nodes.
            // Example update: set display type to Display for top‑level math paragraphs.
            NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
            for (int i = 0; i < officeMathNodes.Count; i++)
            {
                OfficeMath om = (OfficeMath)officeMathNodes[i];

                // Only top‑level math paragraphs have MathObjectType.OMathPara.
                if (om.MathObjectType == MathObjectType.OMathPara)
                    om.DisplayType = OfficeMathDisplayType.Display;
            }

            // Validate each OfficeMath node against the expected type.
            bool allValid = true;
            for (int i = 0; i < officeMathNodes.Count; i++)
            {
                OfficeMath om = (OfficeMath)officeMathNodes[i];
                MathObjectType actual = om.MathObjectType;

                // If an expected type is defined for this index, compare it.
                if (expectedTypes != null && expectedTypes.TryGetValue(i, out MathObjectType expected))
                {
                    if (actual != expected)
                    {
                        allValid = false;
                        Console.WriteLine($"Mismatch at OfficeMath index {i}: expected {expected}, actual {actual}.");
                    }
                }
            }

            // Save the (potentially modified) document (lifecycle rule: save)
            doc.Save(outputPath, SaveFormat.Docx);

            return allValid;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Adjust paths as needed for your environment.
            string input = @"C:\Docs\Source.docx";
            string output = @"C:\Docs\Result.docx";

            var expected = new Dictionary<int, MathObjectType>
            {
                { 0, MathObjectType.OMathPara },   // first OfficeMath should be a paragraph
                { 1, MathObjectType.Fraction }     // second OfficeMath should be a fraction, etc.
            };

            bool result = OfficeMathValidator.Validate(input, output, expected);
            Console.WriteLine($"Validation result: {result}");
        }
    }
}
