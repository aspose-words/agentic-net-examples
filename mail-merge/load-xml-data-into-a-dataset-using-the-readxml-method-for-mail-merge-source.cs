using System;
using System.Data;
using System.IO;
using Aspose.Words;

namespace AsposeWordsMailMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare a simple XML file that will be used as the mail merge data source.
            string xmlContent = @"
<Root>
    <Person>
        <FullName>Thomas Hardy</FullName>
        <Address>120 Hanover Sq., London</Address>
    </Person>
    <Person>
        <FullName>Paolo Accorti</FullName>
        <Address>Via Monte Bianco 34, Torino</Address>
    </Person>
</Root>";
            string xmlPath = Path.Combine(Path.GetTempPath(), "people.xml");
            File.WriteAllText(xmlPath, xmlContent);

            // Load the XML data into a DataSet using ReadXml.
            DataSet dataSet = new DataSet();
            dataSet.ReadXml(xmlPath);

            // Create a new blank document and add mail merge fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define a mail merge region that matches the DataTable name ("Person").
            builder.InsertField("MERGEFIELD TableStart:Person");
            builder.Writeln(); // optional line break

            // Insert the fields that will be populated from the XML data.
            builder.InsertField("MERGEFIELD FullName");
            builder.Write(" - ");
            builder.InsertField("MERGEFIELD Address");
            builder.Writeln(); // optional line break

            // End of the mail merge region.
            builder.InsertField("MERGEFIELD TableEnd:Person");

            // Perform the mail merge using the DataSet as the source.
            doc.MailMerge.ExecuteWithRegions(dataSet);

            // Save the merged document.
            string outputPath = Path.Combine(Path.GetTempPath(), "MergedDocument.docx");
            doc.Save(outputPath);

            // The example finishes without waiting for user input.
        }
    }
}
