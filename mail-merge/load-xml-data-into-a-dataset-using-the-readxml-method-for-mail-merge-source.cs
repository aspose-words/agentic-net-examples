using System;
using System.Data;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple mail‑merge template with a region named "Customer".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin the region.
        builder.InsertField(" MERGEFIELD TableStart:Customer");
        builder.Writeln();

        // Fields that will be filled from the XML data.
        builder.InsertField(" MERGEFIELD FullName");
        builder.Writeln();
        builder.InsertField(" MERGEFIELD Address");
        builder.Writeln();

        // End the region.
        builder.InsertField(" MERGEFIELD TableEnd:Customer");

        // Load XML data into a DataSet using ReadXml.
        const string xml = @"
<Customers>
    <Customer>
        <FullName>Thomas Hardy</FullName>
        <Address>120 Hanover Sq., London</Address>
    </Customer>
    <Customer>
        <FullName>Paolo Accorti</FullName>
        <Address>Via Monte Bianco 34, Torino</Address>
    </Customer>
</Customers>";

        DataSet dataSet = new DataSet();
        using (StringReader sr = new StringReader(xml))
        {
            dataSet.ReadXml(sr);
        }

        // Perform mail merge using the DataSet (the table name must match the region name).
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the merged document.
        doc.Save("MailMergeResult.docx");
    }
}
