using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample XML data.
        string xmlContent = @"
<Report>
    <Section>
        <Title>Fruits</Title>
        <Entry>
            <Name>Apple</Name>
            <Type>Important</Type>
        </Entry>
        <Entry>
            <Name>Banana</Name>
            <Type>Regular</Type>
        </Entry>
        <Entry>
            <Name>Cherry</Name>
            <Type>Important</Type>
        </Entry>
    </Section>
    <Section>
        <Title>Vegetables</Title>
        <Entry>
            <Name>Carrot</Name>
            <Type>Regular</Type>
        </Entry>
        <Entry>
            <Name>Broccoli</Name>
            <Type>Important</Type>
        </Entry>
    </Section>
</Report>";

        // Parse XML into a strongly‑typed model.
        ReportModel model = ParseXmlToModel(xmlContent);

        // Build the template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Outer numbered list for sections.
        builder.Writeln("<<foreach [section in model.Sections]>>");
        builder.Writeln("1. <<[section.Title]>>");

        // Inner numbered list for filtered entries (restart numbering for each section).
        builder.Writeln("   <<restartNum>><<foreach [entry in section.Entry]>>");
        builder.Writeln("      <<if [entry.Type == \"Important\"]>><<[entry.Name]>> <</if>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // Load the template and generate the report.
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        report.Save(outputPath);
    }

    private static ReportModel ParseXmlToModel(string xml)
    {
        XDocument doc = XDocument.Parse(xml);
        ReportModel model = new ReportModel();

        foreach (XElement sectionElem in doc.Root?.Elements("Section") ?? Enumerable.Empty<XElement>())
        {
            Section section = new Section
            {
                Title = (string?)sectionElem.Element("Title") ?? string.Empty,
                Entry = sectionElem.Elements("Entry")
                                   .Select(e => new Entry
                                   {
                                       Name = (string?)e.Element("Name") ?? string.Empty,
                                       Type = (string?)e.Element("Type") ?? string.Empty
                                   })
                                   .ToList()
            };
            model.Sections.Add(section);
        }

        return model;
    }
}

// Data model classes.
public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

public class Section
{
    public string Title { get; set; } = string.Empty;
    public List<Entry> Entry { get; set; } = new();
}

public class Entry
{
    public string Name { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty;
}
