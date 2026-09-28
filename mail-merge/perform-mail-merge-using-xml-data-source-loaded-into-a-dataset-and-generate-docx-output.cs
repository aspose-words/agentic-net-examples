using System;
using System.Data;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a template document with merge fields.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // First line: Dear <<FirstName>> <<LastName>>,
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "«FirstName»");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "«LastName»");
        builder.Writeln(",");

        // Second line: Your order <<OrderID>> has been shipped on <<ShipDate>>.
        builder.Write("Your order ");
        builder.InsertField("MERGEFIELD OrderID", "«OrderID»");
        builder.Write(" has been shipped on ");
        builder.InsertField("MERGEFIELD ShipDate", "«ShipDate»");
        builder.Writeln(".");

        // Third line: Thank you for shopping with us.
        builder.Writeln("Thank you for shopping with us.");

        // XML data source as a string.
        string xml = @"<?xml version='1.0' encoding='utf-8'?>
<Customers>
  <Customer>
    <FirstName>John</FirstName>
    <LastName>Doe</LastName>
    <OrderID>12345</OrderID>
    <ShipDate>2023-08-01</ShipDate>
  </Customer>
  <Customer>
    <FirstName>Jane</FirstName>
    <LastName>Smith</LastName>
    <OrderID>67890</OrderID>
    <ShipDate>2023-08-02</ShipDate>
  </Customer>
</Customers>";

        // Load XML into a DataSet.
        DataSet dataSet = new DataSet();
        using (StringReader sr = new StringReader(xml))
        {
            dataSet.ReadXml(sr);
        }

        // Get the DataTable that contains the customer records.
        DataTable customerTable = dataSet.Tables["Customer"];

        // Perform mail merge for each record in the DataTable.
        // The Execute method will repeat the document for each row.
        template.MailMerge.Execute(customerTable);

        // Save the merged document.
        template.Save("MergedOutput.docx");
    }
}
