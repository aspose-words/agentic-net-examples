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
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        int summationCount = 0;
        foreach (Node node in mathNodes)
        {
            OfficeMath officeMath = (OfficeMath)node;

            // The Unicode character for the summation symbol is U+2211 (∑).
            // Count the equation if its text contains this symbol.
            if (officeMath.GetText().Contains("\u2211"))
                summationCount++;
        }

        Console.WriteLine($"OfficeMath equations containing a summation symbol: {summationCount}");
    }
}
