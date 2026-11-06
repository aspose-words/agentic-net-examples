using System;
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

        // Build a simple 1x1 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Sample cell");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Set the text wrapping style of the table.
        // The available TextWrapping enum values in this version are None and Around.
        // Use Around as the closest alternative to Square.
        table.TextWrapping = TextWrapping.Around;

        // Save the document.
        string outputPath = "TableTextWrapping.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Load the document again to ensure the setting persisted.
        Document loadedDoc = new Document(outputPath);
        Table loadedTable = loadedDoc.FirstSection.Body.Tables[0];
        if (loadedTable.TextWrapping != TextWrapping.Around)
            throw new Exception("Table text wrapping style was not set correctly.");
    }
}
