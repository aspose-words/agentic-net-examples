using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare an output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create sample XML data.
            // -----------------------------------------------------------------
            string xmlContent =
@"<Report>
  <Categories>
    <Category>
      <Name>Fruits</Name>
      <Items>
        <Item>
          <Name>Apple</Name>
          <Price>1.2</Price>
          <Quantity>10</Quantity>
        </Item>
        <Item>
          <Name>Banana</Name>
          <Price>0.8</Price>
          <Quantity>5</Quantity>
        </Item>
      </Items>
    </Category>
    <Category>
      <Name>Vegetables</Name>
      <Items>
        <Item>
          <Name>Carrot</Name>
          <Price>0.5</Price>
          <Quantity>8</Quantity>
        </Item>
        <Item>
          <Name>Broccoli</Name>
          <Price>1.0</Price>
          <Quantity>3</Quantity>
        </Item>
      </Items>
    </Category>
  </Categories>
</Report>";

            string xmlPath = Path.Combine(outputDir, "Data.xml");
            File.WriteAllText(xmlPath, xmlContent);

            // -----------------------------------------------------------------
            // 2. Build a template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Begin iterating over categories.
            builder.Writeln("<<foreach [category in Report.Categories.Category]>>");
            builder.Writeln("Category: <<[category.Name]>>");
            builder.Writeln("Items:");
            // Iterate over items of the current category.
            builder.Writeln("<<foreach [item in category.Items.Item]>>");
            builder.Writeln("- <<[item.Name]>>: <<[item.Price]>> x <<[item.Quantity]>> = <<[item.Price * item.Quantity]>>");
            builder.Writeln("<</foreach>>");
            // Calculate total for the current category using inline LINQ arithmetic.
            builder.Writeln("Total for <<[category.Name]>>: <<[category.Items.Item.Sum(i => i.Price * i.Quantity)]>>");
            builder.Writeln("<</foreach>>");

            // -----------------------------------------------------------------
            // 3. Load the XML data source (ensure the root object is generated).
            // -----------------------------------------------------------------
            var loadOptions = new XmlDataLoadOptions { AlwaysGenerateRootObject = true };
            XmlDataSource dataSource = new XmlDataSource(xmlPath, loadOptions);

            // -----------------------------------------------------------------
            // 4. Build the report.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // The root object name used in the template is "Report".
            engine.BuildReport(template, dataSource, "Report");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            string resultPath = Path.Combine(outputDir, "ReportResult.docx");
            template.Save(resultPath);
        }
    }
}
