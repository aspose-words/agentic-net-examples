using System;
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

        // Add static heading.
        builder.Writeln("Customer Information");
        builder.Writeln();

        // Insert merge fields.
        builder.InsertField("MERGEFIELD Name");
        builder.Writeln();
        builder.InsertField("MERGEFIELD Address");
        builder.Writeln();

        // Define the merge field names that correspond to the data source.
        string[] fieldNames = new string[] { "Name", "Address" };

        // Data for multiple records.
        var records = new (string Name, string Address)[]
        {
            ("John Doe", "123 Main St"),
            ("Jane Smith", "456 Oak Ave"),
            ("Bob Johnson", "789 Pine Rd")
        };

        // Perform a separate mail merge for each record.
        for (int i = 0; i < records.Length; i++)
        {
            // Clone the template so each iteration works with a fresh copy.
            Document outputDoc = (Document)template.Clone(true);

            // Prepare field values for the current record.
            object[] fieldValues = new object[] { records[i].Name, records[i].Address };

            // Execute the simple mail merge.
            outputDoc.MailMerge.Execute(fieldNames, fieldValues);

            // Remove any merge fields that were not used.
            outputDoc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;

            // Define a deterministic output file name.
            string outputFileName = $"Customer_{i + 1}.pdf";

            // Save the merged document as PDF.
            outputDoc.Save(outputFileName, SaveFormat.Pdf);

            // Verify that the file was created.
            if (!File.Exists(outputFileName))
                throw new InvalidOperationException($"Failed to create output file: {outputFileName}");

            // Ensure the PDF is not empty.
            FileInfo info = new FileInfo(outputFileName);
            if (info.Length == 0)
                throw new InvalidOperationException($"Output file is empty: {outputFileName}");
        }
    }
}
