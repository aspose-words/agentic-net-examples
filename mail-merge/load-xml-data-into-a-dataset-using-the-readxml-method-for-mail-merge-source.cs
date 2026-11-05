using System;
using System.Data;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a template document with merge fields and a region for repeating data.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table region named "Person".
        builder.InsertField("MERGEFIELD TableStart:Person", null);
        builder.Writeln(); // New line after the start field.

        // Insert merge fields that correspond to XML element names.
        builder.InsertField("MERGEFIELD FirstName", null);
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", null);
        builder.Writeln(); // New line for each record.

        // End of the region.
        builder.InsertField("MERGEFIELD TableEnd:Person", null);
        builder.Writeln();

        // Prepare XML data.
        string xmlData = @"
<People>
  <Person>
    <FirstName>John</FirstName>
    <LastName>Doe</LastName>
  </Person>
  <Person>
    <FirstName>Jane</FirstName>
    <LastName>Smith</LastName>
  </Person>
</People>";

        // Load XML into a DataSet.
        DataSet dataSet = new DataSet();
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(xmlData)))
        {
            dataSet.ReadXml(stream);
        }

        // Retrieve the DataTable that matches the region name.
        DataTable personTable = dataSet.Tables["Person"];
        if (personTable == null)
            throw new InvalidOperationException("The required DataTable 'Person' was not found in the XML data.");

        // Ensure the table has rows.
        if (personTable.Rows.Count == 0)
            throw new InvalidOperationException("The DataTable 'Person' contains no rows to merge.");

        // Perform region-based mail merge.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;
        doc.MailMerge.ExecuteWithRegions(personTable);

        // Save the merged document.
        const string outputPath = "MergedOutput.docx";
        doc.Save(outputPath);

        // Validate that the merged output contains expected values.
        string mergedText = doc.GetText();
        if (!mergedText.Contains("John") || !mergedText.Contains("Jane"))
            throw new Exception("Mail merge failed: expected values were not found in the output.");

        // Optionally, indicate success (no interactive output required).
        // The program ends here.
    }
}
