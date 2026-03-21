using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        const string inputPath = "Input.docx";
        const string outputPath = "Output.docx";

        Document doc;

        // Try to load the existing document; if it doesn't exist, create a sample one.
        if (System.IO.File.Exists(inputPath))
        {
            doc = new Document(inputPath);
        }
        else
        {
            doc = new Document();

            // Ensure the document has at least one section.
            Section section = doc.FirstSection ?? (Section)doc.AppendChild(new Section(doc));
            Body body = section.Body ?? (Body)section.AppendChild(new Body(doc));

            // Create a simple 2x3 table as a placeholder.
            Table table = new Table(doc);
            body.AppendChild(table);

            // Apply a built‑in table style.
            table.StyleIdentifier = StyleIdentifier.TableGrid;

            // Add rows and cells.
            for (int r = 0; r < 2; r++)
            {
                Row row = new Row(doc);
                table.AppendChild(row);
                for (int c = 0; c < 3; c++)
                {
                    Cell cell = new Cell(doc);
                    cell.AppendChild(new Paragraph(doc));
                    cell.FirstParagraph.AppendChild(new Run(doc, $"R{r + 1}C{c + 1}"));
                    row.AppendChild(cell);
                }
            }
        }

        // Safely retrieve the first table, if any.
        Table firstTable = null;
        if (doc.FirstSection?.Body?.Tables?.Count > 0)
        {
            firstTable = doc.FirstSection.Body.Tables[0];
        }

        if (firstTable == null)
        {
            Console.WriteLine("No tables found in the document.");
        }
        else
        {
            int rowCount = firstTable.Rows.Count;
            int columnCount = firstTable.Rows.Count > 0 ? firstTable.Rows[0].Cells.Count : 0;

            Console.WriteLine($"Rows: {rowCount}, Columns: {columnCount}");
        }

        // Save the (potentially modified) document.
        doc.Save(outputPath);
    }
}
