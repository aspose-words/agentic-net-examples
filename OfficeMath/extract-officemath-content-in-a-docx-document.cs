using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Math;

class ExtractOfficeMath
{
    static void Main()
    {
        // Path to the source DOCX file that contains OfficeMath objects.
        string inputPath = @"C:\Docs\SourceWithMath.docx";

        // Load the document using the provided Document constructor (load rule).
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // Option 1: Iterate through all OfficeMath nodes and write their
        // plain text representation to the console.
        // -----------------------------------------------------------------
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        Console.WriteLine("Extracted OfficeMath objects (plain text):");
        foreach (OfficeMath om in mathNodes)
        {
            // GetText returns the textual representation of the OfficeMath node.
            Console.WriteLine(om.GetText().Trim());
        }

        // -----------------------------------------------------------------
        // Option 2: Save the whole document as a plain‑text file where
        // OfficeMath objects are exported as LaTeX. This uses the provided
        // TxtSaveOptions and Document.Save (save rule).
        // -----------------------------------------------------------------
        string outputPath = @"C:\Docs\ExtractedMath.txt";

        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            // Export OfficeMath as LaTeX markup.
            OfficeMathExportMode = TxtOfficeMathExportMode.Latex
        };

        // Save the document as a TXT file with the specified options.
        doc.Save(outputPath, txtOptions);

        Console.WriteLine($"OfficeMath content saved to: {outputPath}");
    }
}
