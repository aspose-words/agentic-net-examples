using System;
using Aspose.Words;
using Aspose.Words.Math;

class ListOfficeMathEquations
{
    static void Main()
    {
        // Load the DOCX document from the file system.
        // Using the Document(string) constructor as defined in the provided rules.
        Document doc = new Document("InputDocument.docx");

        // Retrieve all OfficeMath nodes in the document (including those in nested structures).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node and output its textual representation.
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            // GetText() returns the equation text including its child elements.
            Console.WriteLine(officeMath.GetText().Trim());
        }
    }
}
