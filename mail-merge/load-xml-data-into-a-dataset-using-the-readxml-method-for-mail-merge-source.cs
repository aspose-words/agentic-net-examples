using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace MailMergeXmlDataSetExample
{
    public class Program
    {
        public static void Main()
        {
            // Define file paths in the current directory.
            string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "Customers.xml");
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedDocument.docx");

            // Create a simple XML file that will serve as the mail‑merge data source.
            File.WriteAllText(xmlPath,
@"<?xml version=""1.0"" encoding=""utf-8""?>
<Customers>
    <Customer>
        <FullName>Thomas Hardy</FullName>
        <Address>120 Hanover Sq., London</Address>
    </Customer>
    <Customer>
        <FullName>Paolo Accorti</FullName>
        <Address>Via Monte Bianco 34, Torino</Address>
    </Customer>
</Customers>");

            // Load the XML data into a DataSet using the ReadXml method.
            DataSet dataSet = new DataSet();
            dataSet.ReadXml(xmlPath);

            // Create a new blank document that contains mail‑merge fields and a region named "Customers".
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Begin the mail‑merge region.
            builder.InsertField(" MERGEFIELD TableStart:Customers ");

            // Insert fields that correspond to the XML element names.
            builder.InsertField(" MERGEFIELD FullName ");
            builder.Writeln(); // New line between fields.
            builder.InsertField(" MERGEFIELD Address ");

            // End the mail‑merge region.
            builder.InsertField(" MERGEFIELD TableEnd:Customers ");

            // Execute the mail merge using the DataSet as the source.
            // The ExecuteWithRegions overload processes the "Customers" region.
            doc.MailMerge.ExecuteWithRegions(dataSet);

            // Save the merged document to disk.
            doc.Save(outputPath);
        }
    }
}
