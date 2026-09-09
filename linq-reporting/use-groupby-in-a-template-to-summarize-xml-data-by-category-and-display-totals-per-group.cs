using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for XML loading on .NET Core.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample XML data file.
        // -----------------------------------------------------------------
        const string xmlPath = "data.xml";
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<persons>
    <person>
        <Category>Food</Category>
        <Amount>10</Amount>
    </person>
    <person>
        <Category>Food</Category>
        <Amount>20</Amount>
    </person>
    <person>
        <Category>Travel</Category>
        <Amount>15</Amount>
    </person>
    <person>
        <Category>Travel</Category>
        <Amount>5</Amount>
    </person>
    <person>
        <Category>Supplies</Category>
        <Amount>12</Amount>
    </person>
</persons>";
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build a template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Title
        builder.Writeln("Summary of amounts by category:");
        builder.Writeln();

        // GroupBy expression inside a foreach loop.
        // The expression groups the collection 'persons' by the 'Category' element.
        builder.Writeln("<<foreach [g in persons.GroupBy(p => p.Category)]>>");
        builder.Writeln("Category: <<[g.Key]>>");
        // Sum the Amount values directly; the engine automatically converts string values to numbers.
        builder.Writeln("Total Amount: <<[g.Sum(p => p.Amount)]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, just to visualize the tags if needed).
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and bind the XML data source.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // -----------------------------------------------------------------
        // 4. Build the report.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "persons");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "Report.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
