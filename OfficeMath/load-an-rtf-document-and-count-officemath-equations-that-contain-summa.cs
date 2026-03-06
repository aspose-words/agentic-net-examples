using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load the RTF document using RtfLoadOptions.
        RtfLoadOptions loadOptions = new RtfLoadOptions();
        Document doc = new Document("input.rtf", loadOptions);

        // Retrieve all OfficeMath nodes in the document.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        int summationCount = 0;
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            // The summation symbol is Unicode U+2211 (∑).
            if (officeMath.GetText().Contains("\u2211"))
                summationCount++;
        }

        Console.WriteLine($"OfficeMath equations containing summation symbols: {summationCount}");
    }
}
