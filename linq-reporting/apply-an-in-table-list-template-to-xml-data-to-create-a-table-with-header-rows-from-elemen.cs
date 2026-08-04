using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for XML parsing (required for .NET Core).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // ---------- Sample XML data ----------
        string xmlContent = @"<Items>
    <Item>
        <Name>Apple</Name>
        <Price>1.20</Price>
    </Item>
    <Item>
        <Name>Banana</Name>
        <Price>0.80</Price>
    </Item>
    <Item>
        <Name>Cherry</Name>
        <Price>2.50</Price>
    </Item>
</Items>";
        string xmlPath = "data.xml";
        File.WriteAllText(xmlPath, xmlContent);

        // ---------- Create template document ----------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Opening foreach tag – placed before the table so the whole table is repeated per item.
        builder.Writeln("<<foreach [item in Items]>>");

        // Start a table.
        Table table = builder.StartTable();

        // Header row – column names taken from the XML element names.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();

        // Data row – will be repeated for each <Item> element.
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Price]>>");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Closing foreach tag.
        builder.Writeln("<</foreach>>");

        // ---------- Load XML data source ----------
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // ---------- Build the report ----------
        ReportingEngine engine = new ReportingEngine();
        // The root object name in the template is "Items".
        engine.BuildReport(template, dataSource, "Items");

        // ---------- Save the generated report ----------
        string outputPath = "Report.docx";
        template.Save(outputPath);
    }
}
