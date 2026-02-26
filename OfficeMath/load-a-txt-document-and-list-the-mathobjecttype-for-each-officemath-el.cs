using System;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Loading;

class ListOfficeMathTypes
{
    static void Main()
    {
        // Load a TXT document. The TxtLoadOptions can be customized if needed.
        var loadOptions = new TxtLoadOptions();
        Document doc = new Document("Input.txt", loadOptions);

        // Retrieve all OfficeMath nodes in the document (including those inside other OfficeMath nodes).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node and output its MathObjectType.
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            Console.WriteLine($"OfficeMath index {officeMath.IndexOf(officeMath)}: {officeMath.MathObjectType}");
        }
    }
}
