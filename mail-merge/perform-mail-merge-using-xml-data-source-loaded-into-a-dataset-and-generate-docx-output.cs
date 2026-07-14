using System;
using System.Data;
using System.IO;
using Aspose.Words;

public class MailMergeFromXml
{
    public static void Main()
    {
        // Create a simple mail merge template with a repeatable region named "Employee".
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin the region.
        builder.InsertField(" MERGEFIELD TableStart:Employee");
        builder.Writeln("Name: ");
        builder.InsertField(" MERGEFIELD Name");
        builder.Writeln();
        builder.Writeln("Address: ");
        builder.InsertField(" MERGEFIELD Address");
        builder.Writeln();
        // End the region.
        builder.InsertField(" MERGEFIELD TableEnd:Employee");

        // XML data source containing employee records.
        string xmlData = @"
<Employees>
    <Employee>
        <Name>John Doe</Name>
        <Address>123 Main St</Address>
    </Employee>
    <Employee>
        <Name>Jane Smith</Name>
        <Address>456 Oak Ave</Address>
    </Employee>
</Employees>";

        // Load the XML into a DataSet.
        DataSet dataSet = new DataSet();
        using (StringReader sr = new StringReader(xmlData))
        {
            dataSet.ReadXml(sr);
        }

        // Ensure the table name matches the mail merge region name.
        if (dataSet.Tables.Count > 0)
        {
            dataSet.Tables[0].TableName = "Employee";
        }

        // Perform mail merge using the DataSet with regions.
        template.MailMerge.ExecuteWithRegions(dataSet);

        // Save the merged document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.docx");
        template.Save(outputPath);
    }
}
