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

        // Insert two tables, each preceded by a caption that uses a SEQ field.
        // The SEQ field will automatically update its number when fields are refreshed.
        for (int i = 0; i < 2; i++)
        {
            // Caption paragraph: "Table {SEQ Table \* ARABIC}: Sample Table"
            builder.Write("Table ");
            builder.InsertField("SEQ Table \\* ARABIC");
            builder.Writeln(": Sample Table");

            // Build a simple 2x2 table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("R1C1");
            builder.InsertCell();
            builder.Write("R1C2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("R2C1");
            builder.InsertCell();
            builder.Write("R2C2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Add a blank paragraph after each table for readability.
            builder.Writeln();
        }

        // Update all fields in the document so that the SEQ numbers reflect the actual count.
        doc.UpdateFields();

        // Save the document to the local file system.
        doc.Save("TableWithCaption.docx");
    }
}
