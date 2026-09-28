using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a template document with merge fields.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Dear <<Name>>,");
        builder.Writeln("Your appointment is scheduled on <<Date>>.");
        builder.Writeln("Thank you.");

        // Prepare data for multiple merges.
        var records = new List<Dictionary<string, object>>
        {
            new Dictionary<string, object>
            {
                { "Name", "Alice" },
                { "Date", new DateTime(2023, 12, 1).ToString("D") }
            },
            new Dictionary<string, object>
            {
                { "Name", "Bob" },
                { "Date", new DateTime(2023, 12, 2).ToString("D") }
            },
            new Dictionary<string, object>
            {
                { "Name", "Charlie" },
                { "Date", new DateTime(2023, 12, 3).ToString("D") }
            }
        };

        // Perform mail merge for each record, cloning the template each time.
        for (int i = 0; i < records.Count; i++)
        {
            // Clone the template to keep it unchanged for the next iteration.
            Document mergedDoc = (Document)template.Clone();

            // Execute mail merge with the current record's values.
            mergedDoc.MailMerge.Execute(
                new[] { "Name", "Date" },
                new object[] { records[i]["Name"], records[i]["Date"] });

            // Save the merged document with a unique file name.
            string outputPath = $"MergedDocument_{i + 1}.docx";
            mergedDoc.Save(outputPath);
        }
    }
}
