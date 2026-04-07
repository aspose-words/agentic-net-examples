using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // ---------- Prepare sample XML data ----------
        const string xmlFile = "data.xml";
        File.WriteAllText(xmlFile,
@"<Root>
  <Groups>
    <Group>
      <Name>Fruits</Name>
      <Items>
        <Item>Apple</Item>
        <Item>Banana</Item>
        <Item>Cherry</Item>
      </Items>
    </Group>
    <Group>
      <Name>Vegetables</Name>
      <Items>
        <Item>Carrot</Item>
        <Item>Tomato</Item>
        <Item>Spinach</Item>
      </Items>
    </Group>
  </Groups>
</Root>");

        // ---------- Create a template document with an in‑table list ----------
        const string templateFile = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Products Report");
        builder.Writeln();

        // Table header.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Category");
        builder.InsertCell();
        builder.Writeln("Items");
        builder.EndRow();

        // Row that will be repeated for each group.
        builder.InsertCell();
        // Outer foreach over groups.
        builder.Writeln("<<foreach [group in Groups.Group]>>");
        builder.Writeln("<<[group.Name]>>");
        builder.InsertCell();

        // Nested foreach over items.
        builder.Writeln("<<foreach [item in group.Items.Item]>>");
        builder.Writeln("- <<[item]>>");
        builder.Writeln("<</foreach>>");

        // Close the outer foreach.
        builder.Writeln("<</foreach>>");
        builder.EndRow();

        // End of table.
        builder.EndTable();

        // ---------- Save the template and reload it ----------
        templateDoc.Save(templateFile);
        var doc = new Document(templateFile);

        // ---------- Load XML data source ----------
        var loadOptions = new XmlDataLoadOptions { AlwaysGenerateRootObject = true };
        var dataSource = new XmlDataSource(xmlFile, loadOptions);

        // ---------- Build the report ----------
        var engine = new ReportingEngine { Options = ReportBuildOptions.None };
        engine.BuildReport(doc, dataSource);

        // ---------- Save the final document ----------
        const string outputFile = "output.docx";
        doc.Save(outputFile);
    }
}
