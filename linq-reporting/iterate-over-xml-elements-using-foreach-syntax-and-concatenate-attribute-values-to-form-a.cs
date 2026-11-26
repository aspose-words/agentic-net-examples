using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Working directory.
        string workDir = Directory.GetCurrentDirectory();

        // 1. Create sample XML data file.
        string xmlPath = Path.Combine(workDir, "data.xml");
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Items>
    <Item Name=""Apple"" Color=""Red"" />
    <Item Name=""Banana"" Color=""Yellow"" />
    <Item Name=""Grape"" Color=""Purple"" />
</Items>";
        File.WriteAllText(xmlPath, xmlContent);

        // 2. Create a Word template with LINQ Reporting tags.
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Item List:");
        // Iterate over the rows of the Items table (each row corresponds to an <Item> element).
        builder.Writeln("<<foreach [item in Items]>>");
        // Concatenate attribute values (Name and Color) into a single line.
        builder.Writeln("<<[item.Name]>> - <<[item.Color]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // 3. Load the template document.
        Document doc = new Document(templatePath);

        // 4. Prepare the XML data source.
        XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

        // 5. Build the report.
        ReportingEngine engine = new ReportingEngine();
        // Root element name is "Items".
        engine.BuildReport(doc, xmlDataSource, "Items");

        // 6. Save the generated report.
        string outputPath = Path.Combine(workDir, "Report.docx");
        doc.Save(outputPath);
    }
}
