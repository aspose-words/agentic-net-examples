using System;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Loading;

class SetOfficeMathInline
{
    static void Main()
    {
        // Path to the source RTF file.
        string inputPath = @"C:\Docs\SourceDocument.rtf";

        // Path where the modified document will be saved.
        string outputPath = @"C:\Docs\ModifiedDocument.rtf";

        // Load the RTF document using RtfLoadOptions as required by the lifecycle rule.
        RtfLoadOptions loadOptions = new RtfLoadOptions();
        Document doc = new Document(inputPath, loadOptions);

        // Iterate through all OfficeMath nodes in the document.
        foreach (OfficeMath officeMath in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            // Set each equation to be displayed inline.
            officeMath.DisplayType = OfficeMathDisplayType.Inline;
        }

        // Save the modified document (preserving the original format).
        doc.Save(outputPath);
    }
}
