using Aspose.Words;
using Aspose.Words.Tables;

// Load the existing DOCM document.
Document doc = new Document("Input.docm");

// Get the first table in the document (adjust index if needed).
Table table = doc.FirstSection.Body.Tables[0];

// Determine how many columns the table has by checking the first row.
int columnCount = table.FirstRow.Cells.Count;

// Create a new row that belongs to the same document.
Row newRow = new Row(doc);

// Add the required number of cells to the new row.
// Each cell receives a simple paragraph with placeholder text.
for (int i = 0; i < columnCount; i++)
{
    Cell cell = new Cell(doc);
    cell.FirstParagraph.AppendChild(new Run(doc, $"New cell {i + 1}"));
    newRow.Cells.Add(cell);
}

// Append the new row to the end of the table.
table.Rows.Add(newRow);

// Save the modified document back to DOCM format.
doc.Save("Output.docm");
