using System;
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
        builder.Writeln("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "");
        builder.Writeln(",");
        builder.Writeln("This is a test mail merge document.");

        // Define field names and values for multiple records.
        string[] fieldNames = new string[] { "FirstName", "LastName" };
        object[][] fieldValues = new object[][]
        {
            new object[] { "John", "Doe" },
            new object[] { "Jane", "Smith" },
            new object[] { "Bob", "Johnson" }
        };

        // Perform a separate merge for each record, cloning the template each time.
        for (int i = 0; i < fieldValues.Length; i++)
        {
            // Clone the template to keep it unchanged for the next iteration.
            Document mergedDoc = (Document)template.Clone();

            // Remove any unused fields after the merge.
            mergedDoc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;

            // Execute the mail merge for the current record.
            mergedDoc.MailMerge.Execute(fieldNames, fieldValues[i]);

            // Validate that the merged document contains the expected values.
            string mergedText = mergedDoc.GetText();
            if (!mergedText.Contains((string)fieldValues[i][0]) || !mergedText.Contains((string)fieldValues[i][1]))
            {
                throw new InvalidOperationException($"Mail merge failed for record {i + 1}.");
            }

            // Save the merged document to a file.
            string outputFileName = $"MergedDocument_{i + 1}.docx";
            mergedDoc.Save(outputFileName);
        }
    }
}
