using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOC template that contains the table.
        Document doc = new Document("Template.docx");

        // Get the first table in the document (adjust index if needed).
        Table table = doc.FirstSection.Body.Tables[0];

        // Iterate through each row and apply a conditional hide based on cell content.
        foreach (Row row in table.Rows)
        {
            // Ensure the row has at least one cell to examine.
            if (row.FirstCell != null)
            {
                // Retrieve the plain text of the first cell (trim line breaks).
                string cellText = row.FirstCell.GetText().Trim();

                // Example condition: hide rows where the first cell contains the word "Hide".
                if (cellText.Equals("Hide", StringComparison.OrdinalIgnoreCase))
                {
                    // Mark the entire row as hidden.
                    row.Hidden = true;
                }
            }
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
