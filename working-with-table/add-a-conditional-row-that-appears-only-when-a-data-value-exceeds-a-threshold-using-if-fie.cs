using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a document variable that will be used in the IF field.
        // Change this value to test the conditional row visibility.
        doc.Variables.Add("Qty", "40"); // Example value exceeding the threshold.

        // Start the table.
        Table table = builder.StartTable();

        // ----- Header row -----
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // ----- Data rows -----
        // Row 1
        builder.InsertCell();
        builder.Write("Apples");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        // Row 2
        builder.InsertCell();
        builder.Write("Bananas");
        builder.InsertCell();
        builder.Write("40");
        builder.EndRow();

        // ----- Conditional row -----
        // This row will display the text "Exceeds threshold" only if the variable Qty > 30.
        builder.InsertCell();

        // Insert an IF field using the strongly‑typed API.
        FieldIf ifField = (FieldIf)builder.InsertField(FieldType.FieldIf, true);
        ifField.LeftExpression = "Qty";
        ifField.ComparisonOperator = ">";
        ifField.RightExpression = "30";
        ifField.TrueText = "Exceeds threshold";
        ifField.FalseText = string.Empty;
        ifField.Update();

        // Second cell (empty) to keep the table structure.
        builder.InsertCell();
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Update all fields so the IF field evaluates with the current variable value.
        doc.UpdateFields();

        // Save the document to a local file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ConditionalRow.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
