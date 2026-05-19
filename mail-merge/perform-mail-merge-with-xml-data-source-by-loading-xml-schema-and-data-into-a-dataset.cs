using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare temporary folder for files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "MailMergeDemo");
        Directory.CreateDirectory(workDir);

        // Create XML schema file.
        string schemaPath = Path.Combine(workDir, "persons.xsd");
        File.WriteAllText(schemaPath,
@"<?xml version=""1.0"" encoding=""utf-8""?>
<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:element name=""persons"">
    <xs:complexType>
      <xs:sequence>
        <xs:element name=""person"" maxOccurs=""unbounded"">
          <xs:complexType>
            <xs:sequence>
              <xs:element name=""FullName"" type=""xs:string""/>
              <xs:element name=""Age"" type=""xs:int""/>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>");

        // Create XML data file.
        string dataPath = Path.Combine(workDir, "persons.xml");
        File.WriteAllText(dataPath,
@"<?xml version=""1.0"" encoding=""utf-8""?>
<persons>
  <person>
    <FullName>John Doe</FullName>
    <Age>30</Age>
  </person>
  <person>
    <FullName>Jane Smith</FullName>
    <Age>25</Age>
  </person>
</persons>");

        // Build a simple mail‑merge template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        // Start a mail‑merge region named "persons".
        builder.InsertField(" MERGEFIELD TableStart:persons");
        builder.Write("\tName: ");
        builder.InsertField(" MERGEFIELD FullName");
        builder.Write("\tAge: ");
        builder.InsertField(" MERGEFIELD Age");
        builder.InsertParagraph();
        // End the region.
        builder.InsertField(" MERGEFIELD TableEnd:persons");

        // Load XML data into a DataSet using the schema.
        DataSet dataSet = new DataSet();
        dataSet.ReadXmlSchema(schemaPath);
        dataSet.ReadXml(dataPath);

        // Perform mail merge with regions using the DataSet.
        template.MailMerge.ExecuteWithRegions(dataSet);

        // Save the result.
        string outputPath = Path.Combine(workDir, "MergedResult.docx");
        template.Save(outputPath);
    }
}
