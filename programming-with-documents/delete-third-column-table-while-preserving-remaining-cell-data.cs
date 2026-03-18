using System;
using Aspose.Words;
using Aspose.Words.Tables;

class DeleteThirdColumn
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("input.docx");

        // Get the first table in the document (adjust index as needed).
        Table table = doc.FirstSection.Body.Tables[0];

        // Iterate through each row and remove the third cell (index 2) if it exists.
        foreach (Row row in table.Rows)
        {
            // Ensure the row has at least three cells.
            if (row.Cells.Count > 2)
            {
                // Remove the cell at index 2, which corresponds to the third column.
                row.Cells.RemoveAt(2);
            }
        }

        // Save the modified document.
        doc.Save("output.docx");
    }
}
