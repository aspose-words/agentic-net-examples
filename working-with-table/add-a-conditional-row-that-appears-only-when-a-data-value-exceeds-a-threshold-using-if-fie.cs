using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with three rows: header, data, and a conditional row.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Amount");
        builder.EndRow();

        // Data row with a MERGEFIELD for the amount value.
        builder.InsertCell();
        builder.Writeln("Sample");
        builder.InsertCell();
        // Insert a merge field named "Amount".
        builder.InsertField("MERGEFIELD Amount \\* MERGEFORMAT", null);
        builder.EndRow();

        // Conditional row that should appear only when Amount > 100.
        // Each cell contains an IF field that evaluates the MERGEFIELD.
        builder.InsertCell();
        // IF field: if Amount > 100 display "High", otherwise display nothing.
        builder.InsertField("IF { MERGEFIELD Amount } > 100 \"High\" \"\"", null);
        builder.InsertCell();
        // Second cell left empty (could contain additional conditional content).
        builder.Writeln(string.Empty);
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Perform mail merge to supply the Amount value.
        doc.MailMerge.Execute(
            new string[] { "Amount" },
            new object[] { 150 } // Change this value to test the condition.
        );

        // Save the document.
        string outputPath = "ConditionalRow.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // The program ends here without waiting for user input.
    }
}
