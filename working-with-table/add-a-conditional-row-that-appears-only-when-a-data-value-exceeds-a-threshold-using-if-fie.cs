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

        // Add a document variable that will be used in the IF field.
        // Change this value to test the conditional row (e.g., 80 will hide the row).
        doc.Variables.Add("Value", "150");

        // Start the table.
        Table table = builder.StartTable();

        // ----- Header row -----
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // ----- Data row -----
        builder.InsertCell();
        builder.Write("Apples");
        builder.InsertCell();
        builder.Write("150");
        builder.EndRow();

        // ----- Conditional row -----
        // The IF field checks the document variable "Value". If it is greater than 100,
        // the text "Value exceeds 100" will be displayed; otherwise the cell will be empty.
        builder.InsertCell();
        // Insert an IF field: IF { DOCVARIABLE Value } > 100 "Value exceeds 100" ""
        builder.InsertField(" IF  { DOCVARIABLE Value } > 100 \"Value exceeds 100\" \"\" ");
        builder.InsertCell(); // Empty second cell to keep table structure.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Update fields so the IF field evaluates before saving.
        doc.UpdateFields();

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ConditionalTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
