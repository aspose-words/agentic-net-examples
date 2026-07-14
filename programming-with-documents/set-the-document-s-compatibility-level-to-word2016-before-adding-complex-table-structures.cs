using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Settings;   // Needed for MsWordVersion enum

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Set the compatibility level to Word 2016.
        // This ensures the document behaves like it was created in Microsoft Word 2016.
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

        // Use DocumentBuilder to construct the outer table.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table outerTable = builder.StartTable();

        // Build a 3x3 outer table with sample text.
        for (int row = 1; row <= 3; row++)
        {
            for (int col = 1; col <= 3; col++)
            {
                builder.InsertCell();
                builder.Write($"Outer R{row}C{col}");
            }
            builder.EndRow();
        }

        // Finish the outer table.
        builder.EndTable();

        // Create a nested table (2x2) and insert it into the first cell of the outer table.
        Table innerTable = new Table(doc);
        for (int i = 1; i <= 2; i++)
        {
            Row innerRow = new Row(doc);
            innerTable.AppendChild(innerRow);
            for (int j = 1; j <= 2; j++)
            {
                Cell innerCell = new Cell(doc);
                innerCell.AppendChild(new Paragraph(doc));
                innerCell.FirstParagraph.AppendChild(new Run(doc, $"Inner R{i}C{j}"));
                innerRow.AppendChild(innerCell);
            }
        }

        // Append the nested table to the first cell of the outer table.
        outerTable.FirstRow.FirstCell.AppendChild(innerTable);

        // Save the document to a file in the current directory.
        doc.Save("ComplexTable.docx");
    }
}
