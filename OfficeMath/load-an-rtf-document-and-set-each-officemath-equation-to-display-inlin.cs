using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Math;

namespace OfficeMathInlineExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source RTF file.
            string inputPath = @"C:\Docs\source.rtf";

            // Path where the modified document will be saved.
            string outputPath = @"C:\Docs\source_inline.rtf";

            // Load the RTF document using RtfLoadOptions as required by the rule set.
            RtfLoadOptions loadOptions = new RtfLoadOptions();
            Document doc = new Document(inputPath, loadOptions);

            // Iterate over all OfficeMath nodes in the document.
            NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
            foreach (OfficeMath officeMath in officeMathNodes)
            {
                // Set each equation to be displayed inline.
                officeMath.DisplayType = OfficeMathDisplayType.Inline;
                // When DisplayType is Inline, Justification must not be set to a non‑inline value.
                // Therefore we leave Justification unchanged (default is Inline).
            }

            // Save the modified document back to RTF format.
            doc.Save(outputPath);
        }
    }
}
