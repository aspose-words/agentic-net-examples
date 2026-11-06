using System;
using System.Data;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert merge fields into the template.
        builder.Writeln("Employee Information:");
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln();

        // Save the template (optional, for debugging).
        const string templatePath = "Template.docx";
        doc.Save(templatePath);

        // XML data source.
        string xmlData = @"
<DataSet>
    <Person>
        <FirstName>John</FirstName>
        <LastName>Doe</LastName>
    </Person>
</DataSet>";

        // Load XML into a DataSet.
        DataSet dataSet = new DataSet();
        using (StringReader sr = new StringReader(xmlData))
        {
            dataSet.ReadXml(sr);
        }

        // Retrieve the DataTable that contains the merge data.
        DataTable table = dataSet.Tables["Person"];
        if (table == null || table.Rows.Count == 0)
            throw new InvalidOperationException("No data found for merge.");

        // Prepare field names and values for the simple mail merge.
        string[] fieldNames = table.Columns.Cast<DataColumn>()
                                          .Select(col => col.ColumnName)
                                          .ToArray();

        object[] fieldValues = table.Rows[0].ItemArray;

        // Execute the mail merge.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document.
        const string outputPath = "MergedOutput.docx";
        doc.Save(outputPath);

        // Validation: ensure the output file exists and is not empty.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
            throw new InvalidOperationException("Merged document was not created correctly.");

        // The program ends here without waiting for user input.
    }
}
