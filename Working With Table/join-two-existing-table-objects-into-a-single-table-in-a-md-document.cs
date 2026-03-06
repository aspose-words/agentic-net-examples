using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source Markdown document that contains at least two tables.
        const string inputPath = @"C:\Docs\Input.md";

        // Load the Markdown document.
        Document doc = new Document(inputPath);

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
        const string outputPath = @"C:\Docs\Output.md";
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        doc.Save(outputPath, saveOptions);
    }
}
