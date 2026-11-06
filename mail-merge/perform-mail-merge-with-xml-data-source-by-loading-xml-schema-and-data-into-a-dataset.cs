using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a template document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert merge fields that correspond to the XML data columns.
        builder.Writeln("Employee Information:");
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln();

        // Define an XML schema (XSD) as a string.
        const string xmlSchema = @"
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
  <xs:element name='Employees'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Employee' maxOccurs='unbounded'>
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

        // Define XML data that conforms to the schema.
        const string xmlData = @"
<Employees>
  <Employee>
    <FirstName>John</FirstName>
    <LastName>Doe</LastName>
  </Employee>
  <Employee>
    <FirstName>Jane</FirstName>
    <LastName>Smith</LastName>
  </Employee>
</Employees>";

        // Load the schema and data into a DataSet.
        DataSet dataSet = new DataSet();
        using (StringReader schemaReader = new StringReader(xmlSchema))
        {
            dataSet.ReadXmlSchema(schemaReader);
        }
        using (StringReader dataReader = new StringReader(xmlData))
        {
            dataSet.ReadXml(dataReader);
        }

        // The DataSet contains a table named "Employee" (derived from the element name).
        // Perform a region-based mail merge for each employee.
        // Define the region start and end fields in the template.
        builder.MoveToDocumentEnd();
        builder.InsertField("MERGEFIELD TableStart:Employees");
        builder.Writeln("Employee: ");
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.InsertField("MERGEFIELD TableEnd:Employees");

        // Execute the mail merge with regions using the DataTable.
        DataTable employeeTable = dataSet.Tables["Employee"];
        if (employeeTable == null)
            throw new InvalidOperationException("Employee table not found in the DataSet.");

        doc.MailMerge.ExecuteWithRegions(employeeTable);

        // Optional cleanup: remove any unused merge fields.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;

        // Save the merged document.
        const string outputPath = "MergedOutput.docx";
        doc.Save(outputPath);

        // Validate that the output file was created and is not empty.
        FileInfo fileInfo = new FileInfo(outputPath);
        if (!fileInfo.Exists || fileInfo.Length == 0)
            throw new InvalidOperationException("Merged output file was not created correctly.");

        // The program finishes without requiring user interaction.
    }
}
