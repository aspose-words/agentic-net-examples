using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class MailMergeWithXmlExample
{
    public static void Main()
    {
        // XML data that defines a list of employees.
        const string xmlData = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Employees>
  <Employee>
    <FirstName>John</FirstName>
    <LastName>Doe</LastName>
    <Title>Software Engineer</Title>
  </Employee>
  <Employee>
    <FirstName>Jane</FirstName>
    <LastName>Smith</LastName>
    <Title>Project Manager</Title>
  </Employee>
</Employees>";

        // Load the XML directly into a DataSet, letting it infer the schema.
        DataSet dataSet = new DataSet();
        using (StringReader dataReader = new StringReader(xmlData))
        {
            dataSet.ReadXml(dataReader, XmlReadMode.InferSchema);
        }

        // Retrieve the DataTable that contains the employee rows.
        DataTable employeeTable = dataSet.Tables["Employee"];
        if (employeeTable == null || employeeTable.Rows.Count == 0)
            throw new InvalidOperationException("No employee data found in the XML source.");

        // Create a template document in memory.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Build the template with merge fields that match the column names.
        builder.Writeln("Employee List:");
        builder.Writeln();
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Write(" - ");
        builder.InsertField("MERGEFIELD Title");
        builder.Writeln();

        // Ensure unused fields are removed after the merge.
        templateDoc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;

        // Execute the mail merge using the DataTable.
        templateDoc.MailMerge.Execute(employeeTable);

        // Save the merged document.
        const string outputPath = "MergedOutput.docx";
        templateDoc.Save(outputPath);

        // Validation: output file must exist and contain data.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The merged document was not created.", outputPath);

        FileInfo fileInfo = new FileInfo(outputPath);
        if (fileInfo.Length == 0)
            throw new InvalidOperationException("The merged document is empty.");

        // Verify that merge fields have been replaced.
        Document mergedDoc = new Document(outputPath);
        string mergedText = mergedDoc.GetText();
        if (mergedText.Contains("«FirstName»") ||
            mergedText.Contains("«LastName»") ||
            mergedText.Contains("«Title»"))
        {
            throw new InvalidOperationException("One or more merge fields were not replaced.");
        }

        // Successful completion – no console output required.
    }
}
