using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    // Deletes a table at the specified index from a DOCX file.
    static void Main()
    {
        // Input and output file paths.
        string inputPath = @"C:\Docs\InputDocument.docx";
        string outputPath = @"C:\Docs\OutputDocument.docx";

        // Index of the table to delete (0‑based).
        int tableIndex = 0;

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Retrieve the table node at the given index.
        Table table = doc.GetChild(NodeType.Table, tableIndex, true) as Table;

        // If the table exists, remove it from its parent.
        if (table != null)
        {
            table.Remove();
        }

        // Save the modified document.
        doc.Save(outputPath);
    }
}
