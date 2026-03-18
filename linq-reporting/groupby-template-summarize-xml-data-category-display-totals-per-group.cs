using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the Word template that contains the GroupBy syntax.
        // The template should have something like:
        // <<groupby [persons.person] by Category>>
        //   Category: <<[Category]>>
        //   Total Salary: <<[Salary]:sum>>
        // <<endgroupby>>
        Document template = new Document("Template.docx");

        // XML data that will be grouped by the <Category> element.
        string xml = @"
<persons>
    <person>
        <Name>John</Name>
        <Category>Employee</Category>
        <Salary>5000</Salary>
    </person>
    <person>
        <Name>Jane</Name>
        <Category>Employee</Category>
        <Salary>6000</Salary>
    </person>
    <person>
        <Name>Bob</Name>
        <Category>Contractor</Category>
        <Salary>3000</Salary>
    </person>
</persons>";

        // Create an XmlDataSource from the XML string.
        using (MemoryStream xmlStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml)))
        {
            // The XmlDataSource class resides directly in Aspose.Words.Reporting namespace.
            XmlDataSource dataSource = new XmlDataSource(xmlStream);

            // Build the report – the engine will process the GroupBy tags,
            // aggregate the Salary values per Category and insert the results.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, dataSource, "persons");
        }

        // Save the populated document.
        template.Save("Result.docx");
    }
}
