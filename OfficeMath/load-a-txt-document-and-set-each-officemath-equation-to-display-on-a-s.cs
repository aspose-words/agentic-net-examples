using System;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source TXT file.
        string inputPath = @"C:\Docs\source.txt";

        // Load the TXT document using TxtLoadOptions (default options are sufficient).
        Document doc = new Document(inputPath, new TxtLoadOptions());

        // Iterate over all OfficeMath nodes in the document.
        foreach (OfficeMath officeMath in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            // Ensure the equation is displayed on its own line.
            officeMath.DisplayType = OfficeMathDisplayType.Display;
        }

        // Save the modified document. Saving as DOCX preserves the display formatting.
        string outputPath = @"C:\Docs\result.docx";
        doc.Save(outputPath);
    }
}
