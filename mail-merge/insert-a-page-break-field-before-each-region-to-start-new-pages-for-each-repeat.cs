using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a PAGE_BREAK field before the region to start a new page for each repeat.
        builder.InsertField("PAGE_BREAK");

        // Insert the start of the mail merge region named "Employees".
        builder.InsertField("MERGEFIELD TableStart:Employees");

        // Insert merge fields that will be repeated for each row in the region.
        builder.Writeln(); // Ensure fields are on a new line.
        builder.InsertField("MERGEFIELD Name");
        builder.Writeln();
        builder.InsertField("MERGEFIELD Title");
        builder.Writeln();

        // Insert the end of the mail merge region.
        builder.InsertField("MERGEFIELD TableEnd:Employees");

        // Prepare the data source for the region.
        DataTable employeeTable = new DataTable("Employees");
        employeeTable.Columns.Add("Name", typeof(string));
        employeeTable.Columns.Add("Title", typeof(string));

        employeeTable.Rows.Add("John Doe", "Manager");
        employeeTable.Rows.Add("Jane Smith", "Developer");
        employeeTable.Rows.Add("Bob Johnson", "Designer");

        // Execute the mail merge with regions.
        doc.MailMerge.ExecuteWithRegions(employeeTable);

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.docx");

        // Save the merged document.
        doc.Save(outputPath);

        // Validation: ensure the output file exists.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output file was not created.");

        // Load the saved document to verify that merge fields were replaced.
        Document resultDoc = new Document(outputPath);
        string resultText = resultDoc.GetText();

        // Check that at least one known value appears in the result.
        if (!resultText.Contains("John Doe") || !resultText.Contains("Jane Smith") || !resultText.Contains("Bob Johnson"))
            throw new InvalidOperationException("Merged content does not contain expected values.");

        // Optionally, you could delete the output file after verification.
        // File.Delete(outputPath);
    }
}
