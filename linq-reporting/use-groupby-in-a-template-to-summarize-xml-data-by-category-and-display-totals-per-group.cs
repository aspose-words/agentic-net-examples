using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingGroupByExample
{
    // Simple data model is not used directly; XML is the data source.
    class Program
    {
        static void Main()
        {
            // Ensure the working directory exists.
            string workDir = Directory.GetCurrentDirectory();

            // 1. Create sample XML data.
            string xmlPath = Path.Combine(workDir, "Data.xml");
            File.WriteAllText(xmlPath,
@"<Items>
    <Item>
        <Category>Food</Category>
        <Amount>10.5</Amount>
    </Item>
    <Item>
        <Category>Food</Category>
        <Amount>5.0</Amount>
    </Item>
    <Item>
        <Category>Books</Category>
        <Amount>12.99</Amount>
    </Item>
    <Item>
        <Category>Books</Category>
        <Amount>7.50</Amount>
    </Item>
    <Item>
        <Category>Electronics</Category>
        <Amount>199.99</Amount>
    </Item>
</Items>");

            // 2. Build the template document programmatically.
            string templatePath = Path.Combine(workDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Report Summary");
            builder.Writeln();
            // LINQ Reporting tag: group items by Category and calculate total per group.
            builder.Writeln("<<foreach [g in Item.GroupBy(i => i.Category)]>>");
            builder.Writeln("Category: <<[g.Key]>>");
            builder.Writeln("Total: <<[g.Sum(i => i.Amount)]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // 3. Load the template for reporting.
            Document reportDoc = new Document(templatePath);

            // 4. Create an XML data source.
            XmlDataSource xmlData = new XmlDataSource(xmlPath);

            // 5. Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this simple scenario.
            engine.BuildReport(reportDoc, xmlData, "Item");

            // 6. Save the generated report.
            string outputPath = Path.Combine(workDir, "Report.docx");
            reportDoc.Save(outputPath);
        }
    }
}
