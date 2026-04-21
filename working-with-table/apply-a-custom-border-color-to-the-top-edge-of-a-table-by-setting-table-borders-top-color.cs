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

        // Start a table and add a single row with two cells.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Apply a custom top border to the whole table.
        // Use Table.SetBorder to set the top border's style, width and color.
        table.SetBorder(BorderType.Top, LineStyle.Single, 2.0, Color.Blue, true);

        // Finish the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "CustomTopBorder.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
