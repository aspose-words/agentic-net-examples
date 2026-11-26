using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a template document with merge fields.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Dear ");
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln(",");
        builder.Writeln("This is a personalized letter.");

        // Prepare data for multiple merges.
        string[] firstNames = { "Alice", "Bob", "Charlie" };
        string[] lastNames = { "Anderson", "Brown", "Clark" };

        // Ensure the output directory exists.
        string outputDir = "Output";
        System.IO.Directory.CreateDirectory(outputDir);

        for (int i = 0; i < firstNames.Length; i++)
        {
            // Clone the template to keep each output independent.
            Document doc = (Document)template.Clone(true);

            // Execute simple mail merge.
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;
            doc.MailMerge.Execute(
                new string[] { "FirstName", "LastName" },
                new object[] { firstNames[i], lastNames[i] });

            // Save the merged document.
            string fileName = System.IO.Path.Combine(outputDir, $"Merged_{i + 1}.docx");
            doc.Save(fileName);

            // Validate that the merged content contains the expected values.
            string text = doc.GetText();
            if (!text.Contains(firstNames[i]) || !text.Contains(lastNames[i]))
                throw new InvalidOperationException($"Merge validation failed for file {fileName}.");

            // Optional: also ensure the file was created.
            if (!System.IO.File.Exists(fileName))
                throw new InvalidOperationException($"Output file not found: {fileName}.");
        }
    }
}
