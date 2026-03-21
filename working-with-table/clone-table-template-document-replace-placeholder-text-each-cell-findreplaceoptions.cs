using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Replacing;

class TableCloneAndReplace
{
    static void Main()
    {
        const string templatePath = "Template.docx";

        // Ensure a template document exists; if not, create one with a simple table.
        Document templateDoc;
        if (File.Exists(templatePath))
        {
            templateDoc = new Document(templatePath);
        }
        else
        {
            templateDoc = new Document();
            templateDoc.EnsureMinimum();

            // Create a table with placeholder text in each cell.
            Table table = new Table(templateDoc);
            // Define a simple 2x2 table.
            for (int i = 0; i < 2; i++)
            {
                Row row = new Row(templateDoc);
                for (int j = 0; j < 2; j++)
                {
                    Cell cell = new Cell(templateDoc);
                    // Insert placeholder text.
                    cell.FirstParagraph.AppendChild(new Run(templateDoc, $"_Placeholder{i}{j}_"));
                    row.Cells.Add(cell);
                }
                table.Rows.Add(row);
            }

            // Add the table to the document.
            templateDoc.FirstSection.Body.AppendChild(table);
            templateDoc.Save(templatePath);
        }

        // Retrieve the first table from the template.
        Table sourceTable = (Table)templateDoc.GetChild(NodeType.Table, 0, true);
        if (sourceTable == null)
        {
            Console.WriteLine("No table found in the template document.");
            return;
        }

        // Deep clone the table so it can be inserted into another document.
        Table clonedTable = (Table)sourceTable.Clone(true);

        // Create a new blank document that will receive the cloned table.
        Document targetDoc = new Document();
        targetDoc.EnsureMinimum();

        // Append the cloned table to the body of the first section.
        targetDoc.FirstSection.Body.AppendChild(clonedTable);

        // Define the placeholder text and its replacement.
        var placeholders = new (string placeholder, string replacement)[]
        {
            ("_Placeholder00_", "John"),
            ("_Placeholder01_", "Doe"),
            ("_Placeholder10_", "123 Main St."),
            ("_Placeholder11_", "Cityville")
        };

        // Iterate over every cell in the cloned table and replace placeholders.
        foreach (Row row in clonedTable.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                foreach (var (placeholder, replacement) in placeholders)
                {
                    FindReplaceOptions options = new FindReplaceOptions
                    {
                        MatchCase = true,
                        FindWholeWordsOnly = true
                    };
                    cell.Range.Replace(placeholder, replacement, options);
                }
            }
        }

        // Save the resulting document.
        targetDoc.Save("Result.docx");
        Console.WriteLine("Result.docx has been created successfully.");
    }
}
