using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply a thick top border (4 points) and a thin bottom border (1 point).
        table.SetBorder(BorderType.Top, LineStyle.Single, 4.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.0, Color.Black, true);

        // Save the document.
        string outputPath = "TableBorders.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");

        // The program ends here without waiting for user input.
    }
}
