using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Register code page provider for XML handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample XML data representing a hierarchical structure.
        // -----------------------------------------------------------------
        const string xmlContent = @"
<categories>
    <category Name='Fruits'>
        <item Name='Apple' />
        <item Name='Banana' />
    </category>
    <category Name='Vegetables'>
        <item Name='Carrot' />
        <item Name='Lettuce' />
    </category>
</categories>";
        const string xmlPath = "data.xml";
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // ---------------------------------------------------------------
        // 2. Build a template document that uses LINQ Reporting tags.
        // ---------------------------------------------------------------
        const string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Create a bulleted list (default template) and apply it to the builder.
        List bulletList = templateDoc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;
        builder.ListFormat.ListLevelNumber = 0; // top‑level bullets

        // Outer loop – iterate over categories.
        builder.Writeln("<<foreach [cat in categories]>>");
        // Category name – level 0 bullet.
        builder.Writeln("<<[cat.Name]>>");

        // Switch to level 1 for inner items.
        builder.ListFormat.ListLevelNumber = 1;

        // Inner loop – iterate over items within a category.
        builder.Writeln("<<foreach [it in cat.item]>>");
        // Item name – level 1 bullet.
        builder.Writeln("<<[it.Name]>>");
        builder.Writeln("<</foreach>>");

        // Return to level 0 after finishing inner loop.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // 3. Load the template and bind the XML data source.
        // ---------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        XmlDataSource xmlData = new XmlDataSource(xmlPath);

        // Build the report. The data source name must match the root tag used in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, xmlData, "categories");

        // ---------------------------------------------------------------
        // 4. Save the generated report.
        // ---------------------------------------------------------------
        const string outputPath = "output.docx";
        reportDoc.Save(outputPath);
    }
}
