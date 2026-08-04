using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Create a sample XML file that will serve as the data source.
        // -----------------------------------------------------------------
        const string xmlPath = "report.xml";
        File.WriteAllText(xmlPath,
@"<Report>
  <Sections>
    <Section>
      <Title>Section 1</Title>
      <Items>
        <Item>Item 1.1</Item>
        <Item>Item 1.2</Item>
      </Items>
    </Section>
    <Section>
      <Title>Section 2</Title>
      <Items>
        <Item>Item 2.1</Item>
      </Items>
    </Section>
  </Sections>
</Report>");

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Top‑level numbered list (1., 2., …).
        List topList = templateDoc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = topList;

        // Begin looping over the Section elements.
        builder.Writeln("<<foreach [sec in report.Sections]>>");
        // Restart numbering for each top‑level item and write the section title.
        builder.Writeln("<<restartNum>><<[sec.Title]>>");

        // Sub‑list for the items belonging to the current section.
        List subList = templateDoc.Lists.Add(ListTemplate.NumberArabicParenthesis);
        builder.ListFormat.List = subList;
        builder.Writeln("<<foreach [itm in sec.Items]>>");
        builder.Writeln("<<[itm]>>");
        builder.Writeln("<</foreach>>");

        // End the outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report using a strongly‑typed model.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        ReportModel model = LoadReportModel(xmlPath);

        var engine = new ReportingEngine();
        // The root object name used in the template tags is "report".
        engine.BuildReport(doc, model, "report");

        // Save the generated report.
        doc.Save("ReportResult.docx");
    }

    // Loads the XML file into a strongly‑typed object graph that matches the template tags.
    private static ReportModel LoadReportModel(string xmlPath)
    {
        var xDoc = XDocument.Load(xmlPath);
        var model = new ReportModel();

        foreach (var xSection in xDoc.Root?.Element("Sections")?.Elements("Section") ?? new XElement[0])
        {
            var section = new SectionModel
            {
                Title = (string?)xSection.Element("Title") ?? string.Empty,
                Items = new List<string>()
            };

            foreach (var xItem in xSection.Element("Items")?.Elements("Item") ?? new XElement[0])
                section.Items.Add((string?)xItem ?? string.Empty);

            model.Sections.Add(section);
        }

        return model;
    }
}

// ---------------------------------------------------------------------
// Public data model classes that match the template expressions.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<SectionModel> Sections { get; set; } = new();
}

public class SectionModel
{
    public string Title { get; set; } = string.Empty;
    public List<string> Items { get; set; } = new();
}
