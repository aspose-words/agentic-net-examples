using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph.
        builder.Writeln("This is a sample paragraph.");

        // Add a table with 2 rows and 2 columns.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        builder.EndTable();

        // Add a bulleted list.
        builder.Writeln(); // Add spacing before the list.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("First bullet item");
        builder.Writeln("Second bullet item");
        builder.Writeln("Third bullet item");
        builder.ListFormat.RemoveNumbers();

        // Save the document as ODT.
        string outputPath = "output.odt";
        doc.Save(outputPath, SaveFormat.Odt);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            // Optionally reload the document to ensure it was saved correctly.
            Document loadedDoc = new Document(outputPath);
        }
    }
}
