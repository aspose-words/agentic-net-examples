using System;
using System.Data;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for temporary XML data file.
        const string xmlPath = "CustomersData.xml";

        // Simple XML representing a list of customers.
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Customers>
  <Customer>
    <Name>John Doe</Name>
    <Address>123 Main St, Springfield</Address>
  </Customer>
  <Customer>
    <Name>Jane Smith</Name>
    <Address>456 Oak Ave, Metropolis</Address>
  </Customer>
</Customers>";

        // Write the XML to disk so it can be loaded into a DataSet.
        File.WriteAllText(xmlPath, xmlContent);

        // Load the XML into a DataSet. The DataSet will contain a table named "Customer".
        DataSet dataSet = new DataSet();
        dataSet.ReadXml(xmlPath);

        // Create a blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a title.
        builder.Writeln("Customer List:");

        // Define a mail‑merge region that matches the table name ("Customer").
        // TableStart and TableEnd fields mark the region boundaries.
        builder.InsertField(" MERGEFIELD TableStart:Customer");
        // Fields inside the region will be filled for each row of the "Customer" table.
        builder.InsertField(" MERGEFIELD Name");
        builder.Write(" - ");
        builder.InsertField(" MERGEFIELD Address");
        builder.InsertParagraph();
        builder.InsertField(" MERGEFIELD TableEnd:Customer");

        // Perform the mail merge using the DataSet. This will repeat the region for each row.
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the result as a DOCX file.
        const string outputPath = "MailMergeResult.docx";
        doc.Save(outputPath);
    }
}
