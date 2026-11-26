using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a template document in memory.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Build the template content with merge fields.
        builder.Writeln("Dear ");
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln(",");

        builder.Writeln("Your order number is ");
        builder.InsertField("MERGEFIELD OrderNumber");
        builder.Writeln(".");

        // Define merge field names.
        string[] fieldNames = new string[] { "FirstName", "LastName", "OrderNumber" };

        // Define data for multiple records.
        List<object[]> records = new List<object[]>
        {
            new object[] { "John", "Doe", "12345" },
            new object[] { "Jane", "Smith", "67890" },
            new object[] { "Alice", "Brown", "54321" }
        };

        // Ensure output directory exists.
        string outputDir = "OutputPdfs";
        Directory.CreateDirectory(outputDir);

        // Perform mail merge for each record and save as a separate PDF.
        for (int i = 0; i < records.Count; i++)
        {
            // Clone the template for the current record.
            Document doc = (Document)template.Clone();

            // Execute mail merge with cleanup of unused fields.
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;
            doc.MailMerge.Execute(fieldNames, records[i]);

            // Define output file path.
            string outputPath = Path.Combine(outputDir, $"MergedDocument_{i + 1}.pdf");

            // Save the merged document as PDF.
            doc.Save(outputPath, SaveFormat.Pdf);

            // Validate that the file was created and is not empty.
            FileInfo fileInfo = new FileInfo(outputPath);
            if (!fileInfo.Exists || fileInfo.Length == 0)
            {
                throw new InvalidOperationException($"Failed to create PDF for record {i + 1}.");
            }
        }
    }
}
