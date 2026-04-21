using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a DocumentBuilder to construct its contents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // First row – default height.
        builder.InsertCell();
        builder.Write("First row, cell 1");
        builder.InsertCell();
        builder.Write("First row, cell 2");
        builder.EndRow();

        // Configure the next row to have an exact height of 20 points.
        builder.RowFormat.Height = 20.0;
        builder.RowFormat.HeightRule = HeightRule.Exactly;

        // Second row – will use the exact height set above.
        builder.InsertCell();
        builder.Write("Second row, cell 1");
        builder.InsertCell();
        builder.Write("Second row, cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "RowHeightExact.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
