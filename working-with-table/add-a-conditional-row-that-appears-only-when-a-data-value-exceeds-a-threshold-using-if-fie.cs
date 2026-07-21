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

        // Start a table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Add some data rows.
        AddDataRow(builder, "Apples", 20);
        AddDataRow(builder, "Bananas", 150);
        AddDataRow(builder, "Carrots", 80);

        // Define a document variable that will be used in the IF field.
        // In a real scenario this could be set via MailMerge or other means.
        doc.Variables.Add("Qty", "150"); // Change this value to test the condition.

        // Conditional row: appears only when Qty > 100.
        builder.InsertCell();
        // IF field: if the variable Qty is greater than 100, display "Special Offer", otherwise display nothing.
        // Use normal string escaping for the double quotes inside the field code.
        builder.InsertField("IF { DOCVARIABLE Qty } > 100 \"Special Offer\" \"\"");
        builder.InsertCell(); // Second cell left empty.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ConditionalRow.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }

    private static void AddDataRow(DocumentBuilder builder, string item, int quantity)
    {
        builder.InsertCell();
        builder.Write(item);
        builder.InsertCell();
        builder.Write(quantity.ToString());
        builder.EndRow();
    }
}
