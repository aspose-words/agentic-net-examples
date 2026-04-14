using System;
using System.Data;
using System.IO;
using System.Text;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple mail‑merge template with a region named "Person".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("List of persons:");
        // Begin the mail‑merge region.
        builder.InsertField(" MERGEFIELD TableStart:Person");
        // Fields that will be filled from the XML data.
        builder.InsertField(" MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField(" MERGEFIELD LastName");
        // End the mail‑merge region.
        builder.InsertField(" MERGEFIELD TableEnd:Person");
        builder.Writeln();

        // XML data representing a collection of persons.
        string xmlData = @"<?xml version='1.0' encoding='utf-8'?>
<Persons>
  <Person>
    <FirstName>John</FirstName>
    <LastName>Doe</LastName>
  </Person>
  <Person>
    <FirstName>Jane</FirstName>
    <LastName>Smith</LastName>
  </Person>
</Persons>";

        // XML schema (XSD) that describes the structure of the XML above.
        string xmlSchema = @"<?xml version='1.0' encoding='utf-8'?>
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
  <xs:element name='Persons'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Person' maxOccurs='unbounded'>
          <xs:complexType>
            <xs:sequence>
              <xs:element name='FirstName' type='xs:string' />
              <xs:element name='LastName' type='xs:string' />
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

        // Load the schema and the XML data into a DataSet.
        DataSet dataSet = new DataSet();
        using (MemoryStream schemaStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlSchema)))
        using (MemoryStream xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlData)))
        {
            dataSet.ReadXmlSchema(schemaStream);
            dataSet.ReadXml(xmlStream);
        }

        // Perform the mail merge using the DataSet (regions are handled automatically).
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the merged document.
        doc.Save("MailMergeOutput.docx");
    }
}
