using System;
using System.Data;
using System.IO;
using Aspose.Words;

public class MailMergeXmlExample
{
    public static void Main()
    {
        // Create a simple mail merge template document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");
        builder.Writeln();

        // Prepare temporary folder for XML files.
        string dataFolder = Path.Combine(Directory.GetCurrentDirectory(), "MailMergeData");
        Directory.CreateDirectory(dataFolder);

        // XML schema defining the structure of the data.
        string schemaPath = Path.Combine(dataFolder, "persons.xsd");
        File.WriteAllText(schemaPath,
@"<?xml version=""1.0"" encoding=""utf-8""?>
<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:element name=""persons"">
    <xs:complexType>
      <xs:sequence>
        <xs:element name=""person"" maxOccurs=""unbounded"">
          <xs:complexType>
            <xs:sequence>
              <xs:element name=""FirstName"" type=""xs:string"" />
              <xs:element name=""LastName"" type=""xs:string"" />
              <xs:element name=""Message"" type=""xs:string"" />
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>");

        // XML data matching the schema.
        string dataPath = Path.Combine(dataFolder, "persons.xml");
        File.WriteAllText(dataPath,
@"<?xml version=""1.0"" encoding=""utf-8""?>
<persons>
  <person>
    <FirstName>John</FirstName>
    <LastName>Doe</LastName>
    <Message>Hello, this is a merged message.</Message>
  </person>
  <person>
    <FirstName>Jane</FirstName>
    <LastName>Smith</LastName>
    <Message>Welcome to the mail merge example.</Message>
  </person>
</persons>");

        // Load XML schema and data into a DataSet.
        DataSet dataSet = new DataSet();
        dataSet.ReadXmlSchema(schemaPath);
        dataSet.ReadXml(dataPath);

        // Perform mail merge using the first table in the DataSet.
        // The table name will be the root element name ("persons") or the first generated table.
        // Using ExecuteWithRegions allows merging multiple records automatically.
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the merged document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedResult.docx");
        doc.Save(outputPath);
    }
}
