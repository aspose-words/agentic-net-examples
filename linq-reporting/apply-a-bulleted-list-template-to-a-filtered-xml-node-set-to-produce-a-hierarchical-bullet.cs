using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data.
        const string xmlFile = "Categories.xml";
        File.WriteAllText(xmlFile,
@"<Categories>
    <Category Name=""Fruits"">
        <Item>Apple</Item>
        <Item>Banana</Item>
        <Item>Cherry</Item>
    </Category>
    <Category Name=""Vegetables"">
        <Item>Carrot</Item>
        <Item>Broccoli</Item>
    </Category>
    <Category Name=""Beverages"">
        <Item>Tea</Item>
        <Item>Coffee</Item>
        <Item>Juice</Item>
    </Category>
</Categories>");

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Create a single‑level bullet list and assign it to the builder.
        List bulletList = template.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;

        // Outer loop – iterate over categories.
        builder.Writeln("<<foreach [cat in Categories]>>");
        // First level bullet (category name).
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("<<[cat.Name]>>");

        // Inner loop – iterate over items within a category.
        builder.Writeln("<<foreach [it in cat.Item]>>");
        // Second level bullet (item name).
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("<<[it]>>");
        builder.Writeln("<</foreach>>"); // End inner foreach.

        builder.Writeln("<</foreach>>"); // End outer foreach.

        // Save the template to disk.
        const string templateFile = "BulletTemplate.docx";
        template.Save(templateFile);

        // Load the template for reporting.
        Document reportDoc = new Document(templateFile);

        // Load XML data source.
        XmlDataSource dataSource = new XmlDataSource(xmlFile);

        // Build the report. The root object name must match the tag used in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, dataSource, "Categories");

        // Save the generated report.
        const string outputFile = "BulletReport.docx";
        reportDoc.Save(outputFile);
    }
}
