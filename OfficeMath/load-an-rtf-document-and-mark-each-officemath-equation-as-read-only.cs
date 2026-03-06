using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Path to the source RTF file.
        string inputPath = @"C:\Docs\SourceDocument.rtf";

        // Path where the modified document will be saved.
        string outputPath = @"C:\Docs\ReadOnlyEquations.rtf";

        // Load the RTF document using RtfLoadOptions.
        RtfLoadOptions loadOptions = new RtfLoadOptions();
        Document doc = new Document(inputPath, loadOptions);

        // Iterate through all OfficeMath objects in the document.
        // (Aspose.Words does not provide a per‑equation read‑only flag,
        //  so we protect the whole document, which makes every equation read‑only.)
        var officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in officeMathNodes)
        {
            // No per‑node read‑only property exists; placeholder for future logic.
            // Currently, we rely on document‑level protection.
        }

        // Protect the entire document as read‑only.
        doc.Protect(ProtectionType.ReadOnly);

        // Save the modified document back to RTF format.
        doc.Save(outputPath);
    }
}
