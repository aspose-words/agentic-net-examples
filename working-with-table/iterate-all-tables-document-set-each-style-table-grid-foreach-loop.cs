using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing Word document.
        Document doc = new Document("input.docx");

        // Iterate over every table in the document (including nested tables).
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Apply the built‑in "Table Grid" style to the current table.
            table.StyleIdentifier = StyleIdentifier.TableGrid;
        }

        // Save the modified document.
        doc.Save("output.docx");
    }
}
