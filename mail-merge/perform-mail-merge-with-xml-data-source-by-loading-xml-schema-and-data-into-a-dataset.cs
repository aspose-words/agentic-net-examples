using System;
using System.Data;
using System.IO;
using Aspose.Words;

namespace MailMergeXmlExample
{
    public class Program
    {
        public static void Main()
        {
            // Define file names for the XML schema and data.
            const string schemaFile = "people.xsd";
            const string dataFile = "people.xml";
            const string outputFile = "MailMergeResult.docx";

            // Write a simple XML Schema (XSD) that defines the structure of the XML data.
            File.WriteAllText(schemaFile,
@"<?xml version=""1.0"" encoding=""utf-8""?>
<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:element name=""persons"">
    <xs:complexType>
      <xs:sequence>
        <xs:element name=""person"" maxOccurs=""unbounded"">
          <xs:complexType>
            <xs:sequence>
              <xs:element name=""Name"" type=""xs:string"" />
              <xs:element name=""Age"" type=""xs:int"" />
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>");

            // Write sample XML data that conforms to the above schema.
            File.WriteAllText(dataFile,
@"<?xml version=""1.0"" encoding=""utf-8""?>
<persons>
  <person>
    <Name>John Doe</Name>
    <Age>30</Age>
  </person>
  <person>
    <Name>Jane Smith</Name>
    <Age>25</Age>
  </person>
  <person>
    <Name>Bob Johnson</Name>
    <Age>40</Age>
  </person>
</persons>");

            // Load the XML schema and data into a DataSet.
            DataSet dataSet = new DataSet();
            dataSet.ReadXmlSchema(schemaFile);
            dataSet.ReadXml(dataFile);

            // The DataTable that contains the rows we want to merge.
            // The table name is taken from the XSD element name ("person").
            DataTable peopleTable = dataSet.Tables["person"];

            // Create a simple Word document with merge fields that match the column names.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a line that will be repeated for each record.
            builder.InsertField("MERGEFIELD Name", "<Name>");
            builder.Write(" is ");
            builder.InsertField("MERGEFIELD Age", "<Age>");
            builder.Writeln(" years old.");
            // Add a page break after each record for clarity (optional).
            builder.InsertBreak(BreakType.PageBreak);

            // Perform the mail merge using the DataTable.
            doc.MailMerge.Execute(peopleTable);

            // Save the merged document.
            doc.Save(outputFile);
        }
    }
}
