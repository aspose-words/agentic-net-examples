using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a simple template document with mail merge fields.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Customer List:");
        builder.Writeln("Name: <<Name>>");
        builder.Writeln("Address: <<Address>>");
        builder.Writeln("--------------------");
        // Save the template to a file.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // XML schema defining the data structure.
        string xmlSchema = @"<?xml version=""1.0""?>
<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:element name=""Customers"">
    <xs:complexType>
      <xs:sequence>
        <xs:element name=""Customer"" maxOccurs=""unbounded"">
          <xs:complexType>
            <xs:sequence>
              <xs:element name=""Name"" type=""xs:string""/>
              <xs:element name=""Address"" type=""xs:string""/>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

        // XML data matching the schema.
        string xmlData = @"<?xml version=""1.0""?>
<Customers>
  <Customer>
    <Name>John Doe</Name>
    <Address>123 Main St</Address>
  </Customer>
  <Customer>
    <Name>Jane Smith</Name>
    <Address>456 Oak Ave</Address>
  </Customer>
</Customers>";

        // Load schema and data into a DataSet.
        DataSet dataSet = new DataSet();
        using (StringReader schemaReader = new StringReader(xmlSchema))
        {
            dataSet.ReadXmlSchema(schemaReader);
        }
        using (StringReader dataReader = new StringReader(xmlData))
        {
            dataSet.ReadXml(dataReader);
        }

        // Load the template document.
        Document doc = new Document(templatePath);

        // Perform mail merge using the DataTable named "Customer".
        DataTable customerTable = dataSet.Tables["Customer"];
        if (customerTable != null)
        {
            doc.MailMerge.Execute(customerTable);
        }

        // Save the merged document.
        const string outputPath = "MergedDocument.docx";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine("Mail merge completed. Output saved to " + outputPath);
    }
}
