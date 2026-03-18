using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableAutoFitExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing document that contains a table.
            // Replace "Input.docx" with the path to your source document.
            Document doc = new Document("Input.docx");

            // Find the first table in the document.
            // If the document may contain multiple tables, iterate over doc.GetChildNodes(NodeType.Table, true).
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            if (table == null)
            {
                Console.WriteLine("No table found in the document.");
                return;
            }

            // Ensure the table is allowed to auto‑fit its cells.
            table.AllowAutoFit = true;

            // Resize columns proportionally to fit the content.
            // This mimics the Word command "AutoFit to Contents".
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Save the modified document.
            // Replace "Output.docx" with the desired output path.
            doc.Save("Output.docx");
        }
    }
}
