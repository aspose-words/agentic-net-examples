using System;
using System.Data;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a simple XML file that will serve as the mail merge data source.
        string xmlPath = "data.xml";
        File.WriteAllText(xmlPath,
@"<Customers>
    <Customer>
        <FullName>Thomas Hardy</FullName>
        <Address>120 Hanover Sq., London</Address>
    </Customer>
    <Customer>
        <FullName>Paolo Accorti</FullName>
        <Address>Via Monte Bianco 34, Torino</Address>
    </Customer>
</Customers>");

        // Load the XML into a DataSet using the ReadXml method.
        DataSet dataSet = new DataSet();
        dataSet.ReadXml(xmlPath);

        // The DataSet now contains a table named "Customer" with columns FullName and Address.
        DataTable table = dataSet.Tables["Customer"];

        // Create a new blank document and add merge fields that correspond to the XML columns.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertField(" MERGEFIELD FullName ");
        builder.InsertParagraph();
        builder.InsertField(" MERGEFIELD Address ");

        // Perform the mail merge using the DataTable extracted from the DataSet.
        doc.MailMerge.Execute(table);

        // Save the merged document.
        string outputPath = "Merged.docx";
        doc.Save(outputPath);
    }
}
