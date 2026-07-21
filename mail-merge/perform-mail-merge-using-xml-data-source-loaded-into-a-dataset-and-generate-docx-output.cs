using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Path for the XML data source and the output document.
        string xmlPath = "employees.xml";
        string outputPath = "MergedDocument.docx";

        // Create a simple XML file that represents a list of employees.
        // The XML structure will be read into a DataSet, which will create a DataTable named "Employee".
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Root>
    <Employee>
        <FirstName>John</FirstName>
        <LastName>Doe</LastName>
        <Address>123 Main St.</Address>
    </Employee>
    <Employee>
        <FirstName>Jane</FirstName>
        <LastName>Smith</LastName>
        <Address>456 Oak Ave.</Address>
    </Employee>
</Root>";
        File.WriteAllText(xmlPath, xmlContent);

        // Load the XML data into a DataSet.
        DataSet dataSet = new DataSet();
        dataSet.ReadXml(xmlPath);

        // Create a mail‑merge template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a mail‑merge region that will repeat for each row in the "Employee" table.
        // The region name must match the DataTable name ("Employee").
        builder.InsertField("MERGEFIELD TableStart:Employee");
        builder.Write("First Name: ");
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write("\tLast Name: ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Write("\tAddress: ");
        builder.InsertField("MERGEFIELD Address");
        builder.InsertParagraph();
        builder.InsertField("MERGEFIELD TableEnd:Employee");

        // Perform the mail merge using the DataSet with regions.
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the merged document to a DOCX file.
        doc.Save(outputPath);
    }
}
