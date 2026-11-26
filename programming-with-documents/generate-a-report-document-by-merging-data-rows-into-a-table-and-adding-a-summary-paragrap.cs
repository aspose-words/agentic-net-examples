using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Sample data rows to merge into the table
        var records = new List<Record>
        {
            new Record { Id = 1, Name = "Alice", Amount = 120.50 },
            new Record { Id = 2, Name = "Bob", Amount = 75.00 },
            new Record { Id = 3, Name = "Charlie", Amount = 200.75 }
        };

        // Create a new document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a title paragraph
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.Writeln("Sales Report");
        builder.Font.ClearFormatting();

        // Start a table with three columns: Id, Name, Amount
        Table table = builder.StartTable();

        // Header row
        InsertCell(builder, "ID", true);
        InsertCell(builder, "Name", true);
        InsertCell(builder, "Amount", true);
        builder.EndRow();

        // Data rows
        foreach (var rec in records)
        {
            InsertCell(builder, rec.Id.ToString(), false);
            InsertCell(builder, rec.Name, false);
            InsertCell(builder, rec.Amount.ToString("F2"), false);
            builder.EndRow();
        }

        builder.EndTable();

        // Add a summary paragraph after the table
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        builder.Font.Size = 12;
        double totalAmount = 0;
        foreach (var rec in records) totalAmount += rec.Amount;
        builder.Writeln($"Total records: {records.Count}");
        builder.Writeln($"Grand total amount: {totalAmount:F2}");

        // Save the report document
        string outputPath = "Report.docx";
        doc.Save(outputPath);

        // Verify the file was saved and can be reopened
        if (File.Exists(outputPath))
        {
            Document loadedDoc = new Document(outputPath);
        }
    }

    // Helper method to insert a cell with optional header formatting
    private static void InsertCell(DocumentBuilder builder, string text, bool isHeader)
    {
        builder.InsertCell();
        if (isHeader)
        {
            builder.Font.Bold = true;
        }
        builder.Writeln(text);
        builder.Font.ClearFormatting();
    }

    // Simple data model for the report rows
    private class Record
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public double Amount { get; set; }
    }
}
