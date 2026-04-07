using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LinqReportingXmlLoopExample
{
    public static void Main()
    {
        // Define file names in the working directory.
        const string templatePath = "Template.docx";
        const string xmlPath = "Data.xml";
        const string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a Word template that contains LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("People List:");

        // The foreach tag iterates over the XML data source named \"xmlData\".
        // Inside the loop we output the Name and Age elements of each Person.
        builder.Writeln("<<foreach [in xmlData]>>");
        builder.Writeln("- <<[Name]>> (Age: <<[Age]>>)");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Create a simple XML file that will be used as the data source.
        // -----------------------------------------------------------------
        string xmlContent =
@"<People>
    <Person>
        <Name>John Doe</Name>
        <Age>30</Age>
    </Person>
    <Person>
        <Name>Jane Smith</Name>
        <Age>25</Age>
    </Person>
    <Person>
        <Name>Bob Johnson</Name>
        <Age>40</Age>
    </Person>
</People>";

        File.WriteAllText(xmlPath, xmlContent);

        // -----------------------------------------------------------------
        // 3. Load the template, bind the XML data source and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // XmlDataSource reads the XML file; the root element is <People>.
        XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

        // The data source name used in the template tags is \"xmlData\".
        ReportingEngine engine = new ReportingEngine();

        // Build the report – the template tags are replaced with actual data.
        engine.BuildReport(reportDoc, xmlDataSource, "xmlData");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}
