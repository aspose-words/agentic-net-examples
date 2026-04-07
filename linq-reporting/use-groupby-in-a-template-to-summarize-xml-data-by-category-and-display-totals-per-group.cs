using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingGroupByExample
{
    public class Program
    {
        public static void Main()
        {
            // Enable code page support for XML loading (required for some cultures).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a working directory.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
            Directory.CreateDirectory(workDir);

            // 1. Write sample XML data to a file.
            string xmlPath = Path.Combine(workDir, "Data.xml");
            string xmlContent =
@"<Items>
    <Item>
        <Category>Food</Category>
        <Amount>12.5</Amount>
    </Item>
    <Item>
        <Category>Electronics</Category>
        <Amount>99.99</Amount>
    </Item>
    <Item>
        <Category>Food</Category>
        <Amount>7.25</Amount>
    </Item>
    <Item>
        <Category>Books</Category>
        <Amount>15.0</Amount>
    </Item>
    <Item>
        <Category>Electronics</Category>
        <Amount>45.0</Amount>
    </Item>
</Items>";
            File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

            // 2. Build a template document that contains LINQ Reporting tags.
            string templatePath = Path.Combine(workDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Category Summary Report");
            builder.Writeln();

            // Correct LINQ Reporting syntax:
            // The data source name is "items". The XML root is a collection of <Item> elements,
            // so we can iterate directly over "items".
            builder.Writeln("<<foreach [g in items.GroupBy(i => i.Category)]>>");
            builder.Writeln("Category: <<[g.Key]>>   Total Amount: <<[g.Sum(i => i.Amount)]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // 3. Load the template for reporting.
            Document reportDoc = new Document(templatePath);

            // 4. Create an XML data source.
            XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

            // 5. Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, xmlDataSource, "items");

            // 6. Save the generated report.
            string outputPath = Path.Combine(workDir, "Report.docx");
            reportDoc.Save(outputPath);
        }
    }
}
