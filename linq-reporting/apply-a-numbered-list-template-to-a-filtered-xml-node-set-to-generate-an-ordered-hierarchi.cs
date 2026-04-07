using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // ---------- 1. Prepare sample XML data ----------
        const string xmlPath = "data.xml";
        const string xmlContent = @"
<Report>
  <Categories>
    <Category>
      <Name>Category A</Name>
      <Items>
        <Item><Title>Item 1</Title></Item>
        <Item><Title>Item 2</Title></Item>
      </Items>
    </Category>
    <Category>
      <Name>Category B</Name>
      <Items>
        <Item><Title>Item 3</Title></Item>
        <Item><Title>Item 4</Title></Item>
      </Items>
    </Category>
  </Categories>
</Report>";
        File.WriteAllText(xmlPath, xmlContent.Trim());

        // ---------- 2. Create the template document ----------
        const string templatePath = "template.docx";
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Create a numbered list and apply it to the builder.
        List numberedList = template.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = numberedList;

        // First level (categories) – level 0.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("<<restartNum>><<foreach [cat in report.Categories.Category]>>");
        builder.Writeln("<<[cat.Name]>>");

        // Second level (items) – level 1.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("<<restartNum>><<foreach [itm in cat.Items.Item]>>");
        // Show only items whose title contains the digit '3'.
        builder.Writeln("   <<if [itm.Title.Contains(\"3\")]>> <<[itm.Title]>> <</if>>");
        builder.Writeln("<</foreach>>");

        // Close the outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template (required by the lifecycle rule) and reload it.
        template.Save(templatePath);
        Document doc = new Document(templatePath);

        // ---------- 3. Load the XML data source ----------
        // Force generation of a root object so that the tag "report" can be used.
        XmlDataLoadOptions loadOptions = new XmlDataLoadOptions { AlwaysGenerateRootObject = true };
        XmlDataSource dataSource = new XmlDataSource(xmlPath, loadOptions);

        // ---------- 4. Build the report ----------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        bool success = engine.BuildReport(doc, dataSource, "report");

        // ---------- 5. Save the generated report ----------
        const string outputPath = "ReportOutput.docx";
        doc.Save(outputPath);

        Console.WriteLine(success ? "Report generated successfully." : "Report generation failed.");
    }
}
