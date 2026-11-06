using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Build a simple 3x3 table.
        // -----------------------------------------------------------------
        builder.StartTable();

        for (int row = 1; row <= 3; row++)
        {
            for (int col = 1; col <= 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row}C{col}");
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Retrieve the created table.
        Table table = doc.FirstSection.Body.Tables[0];

        // -----------------------------------------------------------------
        // Apply thick outer borders (2 points) and thin inner borders (0.5 points).
        // -----------------------------------------------------------------
        // Outer borders
        table.SetBorder(BorderType.Left, LineStyle.Single, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Right, LineStyle.Single, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Top, LineStyle.Single, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 2.0, Color.Black, true);

        // Inner borders
        table.SetBorder(BorderType.Horizontal, LineStyle.Single, 0.5, Color.Black, true);
        table.SetBorder(BorderType.Vertical, LineStyle.Single, 0.5, Color.Black, true);

        // -----------------------------------------------------------------
        // Save the document.
        // -----------------------------------------------------------------
        string outputPath = "TableStyleExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
