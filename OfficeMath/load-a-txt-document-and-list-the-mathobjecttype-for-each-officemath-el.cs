using System;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Load the TXT document using TxtLoadOptions.
        var loadOptions = new TxtLoadOptions();
        Document doc = new Document("input.txt", loadOptions);

        // Retrieve all OfficeMath nodes in the document.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // List the MathObjectType for each OfficeMath element.
        int index = 0;
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            Console.WriteLine($"OfficeMath #{index}: {officeMath.MathObjectType}");
            index++;
        }
    }
}
