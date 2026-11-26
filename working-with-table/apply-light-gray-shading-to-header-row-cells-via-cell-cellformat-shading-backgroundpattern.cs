using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin a table.
        builder.StartTable();

        // Create header row with three cells, applying light gray shading.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();

            // The builder is now inside the newly created cell.
            Cell cell = builder.CurrentParagraph.ParentNode as Cell;
            if (cell != null)
            {
                cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
                cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
            }

            builder.Writeln($"Header {i + 1}");
        }
        // End the header row.
        builder.EndRow();

        // Add a couple of data rows.
        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"Row {row + 1}, Col {col + 1}");
            }
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "TableWithHeaderShading.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // Reopen the document and verify header shading.
        Document loadedDoc = new Document(outputPath);
        Table table = (Table)loadedDoc.GetChild(NodeType.Table, 0, true);
        Row headerRow = table.Rows[0];
        foreach (Cell headerCell in headerRow.Cells)
        {
            if (headerCell.CellFormat.Shading.BackgroundPatternColor.ToArgb() != Color.LightGray.ToArgb())
                throw new InvalidOperationException("Header cell shading is not light gray as expected.");
        }
    }
}
