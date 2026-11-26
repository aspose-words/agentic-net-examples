using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with two rows.
        // First row contains numeric values.
        builder.StartTable();

        builder.InsertCell();
        builder.Write("10");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        // Second row contains formula fields that sum the values above each column.
        builder.InsertCell();
        builder.InsertField("=SUM(ABOVE)", "0");
        builder.InsertCell();
        builder.InsertField("=SUM(ABOVE)", "0");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the initial document.
        string outputPath = "TableWithFields.docx";
        doc.Save(outputPath);

        // Locate the table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Modify the first cell's numeric value (change 10 to 30).
        Cell cellToModify = table.Rows[0].Cells[0];
        cellToModify.RemoveAllChildren(); // Remove existing content.
        Paragraph newParagraph = new Paragraph(doc);
        newParagraph.AppendChild(new Run(doc, "30"));
        cellToModify.AppendChild(newParagraph);

        // Recalculate all fields (the SUM formulas) after the modification.
        doc.UpdateFields();

        // Save the updated document.
        doc.Save(outputPath);
    }
}
