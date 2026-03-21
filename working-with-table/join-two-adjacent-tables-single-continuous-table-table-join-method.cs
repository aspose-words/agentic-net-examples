using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class JoinAdjacentTables
{
    static void Main()
    {
        // Create a new document with two adjacent tables.
        Document doc = new Document();
        Section section = doc.FirstSection ?? (Section)doc.AppendChild(new Section(doc));
        Body body = section.Body ?? (Body)section.AppendChild(new Body(doc));

        // First table with two rows.
        Table firstTable = new Table(doc);
        firstTable.Rows.Add(CreateRow(doc, "A1", "B1"));
        firstTable.Rows.Add(CreateRow(doc, "A2", "B2"));
        body.AppendChild(firstTable);

        // Second table with two rows.
        Table secondTable = new Table(doc);
        secondTable.Rows.Add(CreateRow(doc, "C1", "D1"));
        secondTable.Rows.Add(CreateRow(doc, "C2", "D2"));
        body.AppendChild(secondTable);

        // Ensure there are at least two tables.
        if (doc.FirstSection.Body.Tables.Count < 2)
        {
            Console.WriteLine("The document does not contain two adjacent tables.");
            return;
        }

        // Move all rows from the second table to the first table.
        foreach (Row row in secondTable.Rows)
        {
            // Clone the row to avoid node-parent conflicts.
            firstTable.Rows.Add(row.Clone(true));
        }

        // Remove the now empty second table.
        secondTable.Remove();

        // Save the modified document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TablesJoined.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Joined table saved to: {outputPath}");
    }

    // Helper method to create a table row with two cells containing the specified texts.
    private static Row CreateRow(Document doc, string cell1Text, string cell2Text)
    {
        Row row = new Row(doc);
        Cell cell1 = new Cell(doc);
        cell1.AppendChild(new Paragraph(doc));
        cell1.FirstParagraph.AppendChild(new Run(doc, cell1Text));
        row.AppendChild(cell1);

        Cell cell2 = new Cell(doc);
        cell2.AppendChild(new Paragraph(doc));
        cell2.FirstParagraph.AppendChild(new Run(doc, cell2Text));
        row.AppendChild(cell2);

        return row;
    }
}
