using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a template document with merge fields.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Customer Information:");
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln();

        // Prepare XML data.
        string xmlContent = @"
<Root>
    <Person>
        <FirstName>John</FirstName>
        <LastName>Doe</LastName>
    </Person>
</Root>";
        // Load XML into a DataSet.
        DataSet dataSet = new DataSet();
        using (MemoryStream ms = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xmlContent)))
        {
            dataSet.ReadXml(ms);
        }

        // Retrieve the DataTable that matches the merge fields.
        if (!dataSet.Tables.Contains("Person"))
            throw new Exception("Expected table 'Person' not found in XML data.");

        DataTable personTable = dataSet.Tables["Person"];
        if (personTable.Rows.Count == 0)
            throw new Exception("No data rows found for merge.");

        // Perform mail merge using the DataTable.
        template.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;
        template.MailMerge.Execute(personTable);

        // Save merged document.
        string outputPath = "MergedOutput.docx";
        template.Save(outputPath);

        // Validate that the merged output contains the expected values.
        Document resultDoc = new Document(outputPath);
        string resultText = resultDoc.GetText();

        if (string.IsNullOrWhiteSpace(resultText) ||
            !resultText.Contains("John") ||
            !resultText.Contains("Doe"))
        {
            throw new Exception("Mail merge failed: merged content does not contain expected values.");
        }

        // Indicate success (no interactive output required, but we can write to console).
        Console.WriteLine("Mail merge completed successfully. Output file: " + Path.GetFullPath(outputPath));
    }
}
