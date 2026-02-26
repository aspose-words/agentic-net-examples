using System;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Path to the DOT (Word template) file.
        // Replace {InputFilePath} with the actual file location.
        string inputPath = "{InputFilePath}";

        // Load the template document.
        Document doc = new Document(inputPath);

        // Retrieve all OfficeMath nodes in the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Assume all equations are inline until proven otherwise.
        bool allInline = true;

        foreach (OfficeMath math in mathNodes)
        {
            // For top‑level OfficeMath objects the DisplayType reflects the layout.
            // Nested OfficeMath objects are always inline, but we still check the property.
            if (math.DisplayType != OfficeMathDisplayType.Inline)
            {
                allInline = false;
                // Output details of the first non‑inline equation found.
                Console.WriteLine($"OfficeMath at index {math.IndexOf(math)} is not inline. DisplayType = {math.DisplayType}");
                break;
            }
        }

        // Report the overall result.
        Console.WriteLine(allInline
            ? "All OfficeMath equations are inline."
            : "Some OfficeMath equations are not inline.");
    }
}
