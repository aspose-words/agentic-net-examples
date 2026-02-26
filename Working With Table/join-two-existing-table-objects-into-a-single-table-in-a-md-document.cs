using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

namespace AsposeWordsTableMerge
{
    class Program
    {
        static void Main()
        {
            // Load the Markdown document that already contains at least two tables.
            Document doc = new Document("input.md");

            // Retrieve the first table from the document's body.
            Table firstTable = doc.FirstSection.Body.Tables[0];

            // Retrieve the second table using the generic GetChild method.
            Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

            // Append all rows from the second table to the first table.
            while (secondTable.HasChildNodes)
                firstTable.Rows.Add(secondTable.FirstRow);

            // Remove the now‑empty second table container.
            secondTable.Remove();

            // Save the modified document back to Markdown format.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
            doc.Save("output.md", saveOptions);
        }
    }
}
