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
            // Paths for temporary files.
            string xmlPath = "people.xml";
            string xsdPath = "people.xsd";
            string outputPath = "Merged.docx";

            // Create a simple XML schema (XSD) describing a list of persons.
            File.WriteAllText(xsdPath,
@"<?xml version=""1.0"" encoding=""utf-8""?>
<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:element name=""Persons"">
    <xs:complexType>
      <xs:sequence>
        <xs:element name=""Person"" maxOccurs=""unbounded"">
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

            // Create a matching XML data file.
            File.WriteAllText(xmlPath,
@"<?xml version=""1.0"" encoding=""utf-8""?>
<Persons>
  <Person>
    <Name>John Doe</Name>
    <Age>30</Age>
  </Person>
  <Person>
    <Name>Jane Smith</Name>
    <Age>25</Age>
  </Person>
</Persons>");

            // Load the XML schema and data into a DataSet.
            DataSet dataSet = new DataSet();
            dataSet.ReadXmlSchema(xsdPath);   // Load schema first.
            dataSet.ReadXml(xmlPath);         // Then load data.

            // Build a simple mail‑merge template document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert merge fields that correspond to the column names in the DataSet.
            builder.InsertField("MERGEFIELD Name", "<Name>");
            builder.Write(" - ");
            builder.InsertField("MERGEFIELD Age", "<Age>");
            builder.Writeln();

            // Perform the mail merge using the first (and only) table from the DataSet.
            doc.MailMerge.Execute(dataSet.Tables[0]);

            // Save the merged document.
            doc.Save(outputPath);
        }
    }
}
