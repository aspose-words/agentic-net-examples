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

        // Move the builder cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Start a table inside the header.
        Table table = builder.StartTable();

        // First row with two cells.
        builder.InsertCell();
        builder.Write("Header Cell 1");
        builder.InsertCell();
        builder.Write("Header Cell 2");
        builder.EndRow();

        // Second row with two cells.
        builder.InsertCell();
        builder.Write("Header Cell 3");
        builder.InsertCell();
        builder.Write("Header Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Return to the main document body and add some regular content.
        builder.MoveToSection(0);
        builder.Writeln("This is the main document body.");

        // Save the document to a file.
        const string outputPath = "HeaderTable.docx";
        doc.Save(outputPath);
    }
}
