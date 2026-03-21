using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing document if it exists; otherwise create a sample document with a table.
        Document doc;
        const string sourceFileName = "Spec.docx";

        if (File.Exists(sourceFileName))
        {
            doc = new Document(sourceFileName);
        }
        else
        {
            doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a simple 2x2 table with numeric values.
            Table table = builder.StartTable();
            builder.InsertCell(); builder.Write("0.5");
            builder.InsertCell(); builder.Write("0.75");
            builder.EndRow();
            builder.InsertCell(); builder.Write("1.25");
            builder.InsertCell(); builder.Write("2");
            builder.EndRow();
            builder.EndTable();
        }

        // Builder used for inserting fields.
        DocumentBuilder fieldBuilder = new DocumentBuilder(doc);

        // Iterate through all tables in the document.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Iterate through each cell of the current table.
            foreach (Cell cell in table.GetChildNodes(NodeType.Cell, true))
            {
                // Get the plain text of the cell (trimmed to ignore whitespace).
                string cellText = cell.GetText().Trim();

                // Try to parse the cell text as a double (e.g., "0.5", "0.75").
                if (double.TryParse(cellText, NumberStyles.Any, CultureInfo.InvariantCulture, out double numericValue))
                {
                    // Remove existing content.
                    cell.RemoveAllChildren();

                    // Ensure the cell has a paragraph to host the field.
                    Paragraph para = new Paragraph(doc);
                    cell.AppendChild(para);

                    // Move the builder cursor to the new paragraph.
                    fieldBuilder.MoveTo(para);

                    // Insert a field that evaluates the numeric value and formats it as a fraction.
                    // The format code "# ?/?" displays the result as a common fraction (e.g., 1/2, 3/4).
                    Field field = fieldBuilder.InsertField($"= {numericValue} \\# \"# ?/?\"");

                    // Update the field so the result is calculated immediately.
                    field.Update();
                }
            }
        }

        // Save the modified document.
        doc.Save("Spec_WithFractions.docx");
    }
}
