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

        // Start a table.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Value");
        builder.EndRow();

        // Data row with a MERGEFIELD for the value.
        builder.InsertCell();
        builder.Write("Sample");
        builder.InsertCell();
        builder.InsertField("MERGEFIELD", "Value");
        builder.EndRow();

        // Conditional row that appears only when the value exceeds 100.
        // The IF field will output the text "High" when the condition is true,
        // otherwise it will output an empty string, effectively hiding the content.
        builder.InsertCell();
        // Insert IF field: IF { MERGEFIELD Value } > 100 "High" ""
        builder.InsertField("IF", "MERGEFIELD Value > 100 \"High\" \"\"");
        builder.InsertCell(); // Second cell can be left empty.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document before mail merge (optional).
        string outputPath = "ConditionalTable.docx";
        doc.Save(outputPath);

        // Perform mail merge with a value that exceeds the threshold.
        doc.MailMerge.Execute(new string[] { "Value" }, new object[] { 150 });

        // Update fields to evaluate the IF condition.
        doc.UpdateFields();

        // Save the final document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
