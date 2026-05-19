using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a simple mail merge source document with a repeatable region named "Employees".
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Employee List:");
            // Begin the mail merge region.
            builder.InsertField(" MERGEFIELD TableStart:Employees");
            // Fields inside the region.
            builder.InsertField(" MERGEFIELD FirstName");
            builder.Write(" ");
            builder.InsertField(" MERGEFIELD LastName");
            builder.Write(" - ");
            builder.InsertField(" MERGEFIELD Title");
            builder.InsertParagraph();
            // End the mail merge region.
            builder.InsertField(" MERGEFIELD TableEnd:Employees");

            // Save the source template (optional, useful for debugging).
            string templatePath = Path.Combine(outputDir, "Template.docx");
            doc.Save(templatePath);

            // Create an XML file that represents the data source.
            string xmlPath = Path.Combine(outputDir, "EmployeesData.xml");
            string xmlContent =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<DataSet>
  <Employees>
    <FirstName>John</FirstName>
    <LastName>Doe</LastName>
    <Title>Manager</Title>
  </Employees>
  <Employees>
    <FirstName>Jane</FirstName>
    <LastName>Smith</LastName>
    <Title>Developer</Title>
  </Employees>
  <Employees>
    <FirstName>Bob</FirstName>
    <LastName>Johnson</LastName>
    <Title>Designer</Title>
  </Employees>
</DataSet>";
            File.WriteAllText(xmlPath, xmlContent);

            // Load the XML into a DataSet.
            DataSet dataSet = new DataSet();
            dataSet.ReadXml(xmlPath);

            // Perform mail merge using the DataSet with regions.
            doc.MailMerge.ExecuteWithRegions(dataSet);

            // Save the merged document.
            string resultPath = Path.Combine(outputDir, "MergedEmployees.docx");
            doc.Save(resultPath);
        }
    }
}
