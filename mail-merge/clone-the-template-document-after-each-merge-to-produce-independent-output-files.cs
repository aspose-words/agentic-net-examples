using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a template document in memory.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Build simple content with merge fields.
        builder.Writeln("Hello ");
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln();
        builder.Writeln("Address: ");
        builder.InsertField("MERGEFIELD Address");
        builder.Writeln();

        // Define merge field names (must match the fields inserted above).
        string[] fieldNames = new[] { "FirstName", "LastName", "Address" };

        // Prepare multiple data records.
        List<object[]> records = new List<object[]>
        {
            new object[] { "John", "Doe", "123 Main St" },
            new object[] { "Jane", "Smith", "456 Oak Ave" },
            new object[] { "Bob", "Johnson", "789 Pine Rd" }
        };

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutputs");
        Directory.CreateDirectory(outputDir);

        // Process each record: clone the template, perform mail merge, and save.
        for (int i = 0; i < records.Count; i++)
        {
            // Clone the template to obtain an independent document for this record.
            Document outputDoc = (Document)template.Clone(true);

            // Execute the mail merge with the current record's values.
            outputDoc.MailMerge.Execute(fieldNames, records[i]);

            // Clean up any unused merge fields.
            outputDoc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;

            // Validate that all merge fields were replaced.
            string mergedText = outputDoc.GetText();
            foreach (object value in records[i])
            {
                if (value == null) continue;
                string valStr = value.ToString();
                if (!mergedText.Contains(valStr))
                {
                    throw new InvalidOperationException($"Merged value '{valStr}' was not found in the document.");
                }
            }

            // Save the merged document to a distinct file.
            string outputPath = Path.Combine(outputDir, $"MergedDocument_{i + 1}.docx");
            outputDoc.Save(outputPath);
        }
    }
}
