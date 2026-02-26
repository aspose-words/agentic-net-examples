using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Retrieve the first table in the document (if any).
        Table table = doc.GetChild(NodeType.Table, 0, true) as Table;

        // Remove the table from its parent node.
        if (table != null)
        {
            table.Remove();
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
