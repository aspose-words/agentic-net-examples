using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        builder.StartTable();

        // Build 3 rows with 4 columns each.
        for (int row = 1; row <= 3; row++)
        {
            for (int col = 1; col <= 4; col++)
            {
                builder.InsertCell();
                builder.Writeln($"Row {row}, Column {col}");
            }
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        string fileName = "TableExample.docx";
        doc.Save(fileName);

        // Verify the file was created.
        if (!File.Exists(fileName))
        {
            throw new InvalidOperationException($"Failed to create {fileName}");
        }
    }
}
