using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Replacing;

class TableCloneAndReplace
{
    static void Main()
    {
        // Load the template document that contains the source table.
        Document templateDoc = new Document("Template.docx");

        // Retrieve the first table from the template.
        Table sourceTable = (Table)templateDoc.GetChild(NodeType.Table, 0, true);

        // Deep clone the table so it can be inserted into another document.
        Table clonedTable = (Table)sourceTable.Clone(true);

        // Create a new blank document that will receive the cloned table.
        Document targetDoc = new Document();
        // Ensure the document has at least one section and one paragraph.
        targetDoc.EnsureMinimum();

        // Append the cloned table to the body of the first section.
        targetDoc.FirstSection.Body.AppendChild(clonedTable);

        // Define the placeholder text and its replacement.
        // In a real scenario these could come from any data source.
        var placeholders = new (string placeholder, string replacement)[]
        {
            ("_FirstName_", "John"),
            ("_LastName_", "Doe"),
            ("_Address_", "123 Main St.")
        };

        // Iterate over every cell in the cloned table and replace placeholders.
        foreach (Row row in clonedTable.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                foreach (var (placeholder, replacement) in placeholders)
                {
                    // Configure find/replace options (case‑sensitive, whole word).
                    FindReplaceOptions options = new FindReplaceOptions
                    {
                        MatchCase = true,
                        FindWholeWordsOnly = true
                    };

                    // Perform the replacement within the cell's range.
                    cell.Range.Replace(placeholder, replacement, options);
                }
            }
        }

        // Save the resulting document.
        targetDoc.Save("Result.docx");
    }
}
