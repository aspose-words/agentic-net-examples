using System;
using Aspose.Words;
using Aspose.Words.Math;

namespace OfficeMathPrinter
{
    class Program
    {
        static void Main()
        {
            // Load the DOCX document that contains OfficeMath objects.
            // (Assumes the file "OfficeMath.docx" exists in the same folder as the executable.)
            Document doc = new Document("OfficeMath.docx");

            // Retrieve all OfficeMath nodes in the document.
            NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

            // Print the textual representation of each OfficeMath object.
            foreach (OfficeMath officeMath in officeMathNodes)
            {
                // GetText returns the plain text of the OfficeMath node and its children.
                Console.WriteLine(officeMath.GetText());
            }
        }
    }
}
