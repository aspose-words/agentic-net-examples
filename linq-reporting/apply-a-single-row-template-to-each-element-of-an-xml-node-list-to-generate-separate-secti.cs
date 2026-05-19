using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

public partial class Program
{
    public static void Main()
    {
        // Create a deterministic output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string xmlPath = Path.Combine(outputDir, "Items.xml");
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string resultPath = Path.Combine(outputDir, "Report.docx");

        // 1. Generate sample XML data.
        CreateSampleXml(xmlPath);

        // 2. Build the LINQ Reporting template programmatically.
        CreateTemplateDocument(templatePath);

        // 3. Load the template and the XML data source.
        Document templateDoc = new Document(templatePath);
        XmlDataSource xmlData = new XmlDataSource(xmlPath);

        // 4. Execute the report. The root name used in the template is "Items".
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(templateDoc, xmlData, "Items");

        // 5. Save the generated report.
        templateDoc.Save(resultPath, SaveFormat.Docx);

        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}.");
        Console.WriteLine($"Result saved to: {resultPath}");
    }

    // Generates a simple XML file with a list of items.
    private static void CreateSampleXml(string filePath)
    {
        string xmlContent =
@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Items>
    <Item>
        <Name>Apple</Name>
        <Price>1.23</Price>
    </Item>
    <Item>
        <Name>Banana</Name>
        <Price>0.99</Price>
    </Item>
    <Item>
        <Name>Cherry</Name>
        <Price>2.50</Price>
    </Item>
</Items>";
        File.WriteAllText(filePath, xmlContent, Encoding.UTF8);
    }

    // Creates a Word template containing a foreach block that repeats a section per XML item.
    private static void CreateTemplateDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title for the whole document.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        builder.Writeln("Product Catalog");

        // Start of the foreach block – iterate over Item elements.
        builder.Writeln("<<foreach [item in Item]>>");

        // Each item will start on a new page.
        builder.InsertBreak(BreakType.PageBreak);

        // Heading for the item.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("<<[item.Name]>>");

        // Simple paragraph with the price.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Price: $<<[item.Price]>>");

        // End of the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath, SaveFormat.Docx);
    }
}
