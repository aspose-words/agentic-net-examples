using System;
using System.Data;
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

        // Add a title.
        builder.Writeln("Employee List:");

        // Insert the start of the mail merge region for a table named "Employees".
        builder.InsertField("MERGEFIELD TableStart:Employees", null);

        // Build a table with merge fields for each column.
        builder.StartTable();

        // First cell – Name field.
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Name", null);

        // Second cell – Age field.
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Age", null);

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Insert the end of the mail merge region.
        builder.InsertField("MERGEFIELD TableEnd:Employees", null);

        // Prepare the data source as a DataTable.
        DataTable employeeTable = new DataTable("Employees");
        employeeTable.Columns.Add("Name", typeof(string));
        employeeTable.Columns.Add("Age", typeof(int));

        // Add sample rows.
        employeeTable.Rows.Add("John Doe", 30);
        employeeTable.Rows.Add("Jane Smith", 28);
        employeeTable.Rows.Add("Bob Johnson", 45);

        // Execute the region-based mail merge.
        doc.MailMerge.ExecuteWithRegions(employeeTable);

        // Save the merged document.
        string outputPath = "MergedOutput.docx";
        doc.Save(outputPath);

        // Validation: ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The merged document was not created.");

        // Load the saved document to verify that merge fields were replaced.
        Document mergedDoc = new Document(outputPath);
        string mergedText = mergedDoc.GetText();

        // Check that at least one known value appears in the output.
        if (!mergedText.Contains("John Doe") || !mergedText.Contains("Jane Smith") || !mergedText.Contains("Bob Johnson"))
            throw new InvalidOperationException("Mail merge did not replace fields as expected.");

        // Ensure no unresolved merge fields remain (they appear as «field» in the text).
        if (mergedText.Contains("«"))
            throw new InvalidOperationException("Unresolved merge fields remain in the document.");

        // The program finishes successfully without user interaction.
    }
}
