using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the source document that contains the specifications table
        Document doc = new Document("Spec.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Iterate through all tables in the document
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Iterate through each cell of the current table
            foreach (Cell cell in table.GetChildNodes(NodeType.Cell, true))
            {
                // Get the plain text of the cell (trimmed to ignore whitespace)
                string cellText = cell.GetText().Trim();

                // Try to parse the cell text as a double (e.g., "0.5", "0.75")
                if (double.TryParse(cellText, NumberStyles.Any, CultureInfo.InvariantCulture, out double numericValue))
                {
                    // Move the builder cursor to the first paragraph of the cell
                    builder.MoveTo(cell.FirstParagraph);

                    // Remove existing runs (original numeric text)
                    cell.RemoveAllChildren();

                    // Insert a field that evaluates the numeric value and formats it as a fraction
                    // The format code "# ?/?" displays the result as a common fraction (e.g., 1/2, 3/4)
                    Field field = builder.InsertField($"= {numericValue} \\# \"# ?/?\"");

                    // Update the field so the result is calculated immediately
                    field.Update();
                }
            }
        }

        // Save the modified document
        doc.Save("Spec_WithFractions.docx");
    }
}
