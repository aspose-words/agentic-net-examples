using System;
using System.IO;
using Aspose.Words;

public class MailMergeExample
{
    public static void Main()
    {
        // Create a template document in memory.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert merge fields into the template.
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" "); // Add a space between fields.
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln(); // End the line.

        // Define field names and data for simple mail merge.
        string[] fieldNames = new string[] { "FirstName", "LastName" };
        object[][] records = new object[][]
        {
            new object[] { "John", "Doe" },
            new object[] { "Jane", "Smith" },
            new object[] { "Alice", "Johnson" }
        };

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "MergedDocs");
        Directory.CreateDirectory(outputDir);

        // Perform mail merge for each record, cloning the template each time.
        for (int i = 0; i < records.Length; i++)
        {
            // Clone the template to keep it unchanged for the next iteration.
            Document mergedDoc = (Document)template.Clone(true);

            // Execute the simple mail merge.
            mergedDoc.MailMerge.Execute(fieldNames, records[i]);

            // Validate that the merge replaced the fields.
            string mergedText = mergedDoc.GetText();
            foreach (object value in records[i])
            {
                if (value == null || !mergedText.Contains(value.ToString()))
                {
                    throw new InvalidOperationException($"Merge failed for record {i}: value '{value}' not found in output.");
                }
            }

            // Save the merged document to a separate file.
            string outputPath = Path.Combine(outputDir, $"MergedDocument_{i + 1}.docx");
            mergedDoc.Save(outputPath);
        }
    }
}
