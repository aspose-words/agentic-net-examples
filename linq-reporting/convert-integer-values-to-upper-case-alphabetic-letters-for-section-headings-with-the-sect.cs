using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class SectionItem
{
    public int Number { get; set; }

    // Returns the uppercase alphabetic representation of the Number (1 -> A, 2 -> B, etc.).
    public string Letter => Number > 0 ? ((char)('A' + Number - 1)).ToString() : string.Empty;

    public string Title { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<SectionItem> Sections { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Create a template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Begin a foreach loop over the Sections collection.
        builder.Writeln("<<foreach [sec in Sections]>>");

        // Use the custom Letter property for alphabetic headings.
        builder.Writeln("<<[sec.Letter]>>. <<[sec.Title]>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // 2. Load the template for reporting.
        var doc = new Document(templatePath);

        // 3. Prepare sample data.
        var model = new ReportModel();
        model.Sections.Add(new SectionItem { Number = 1, Title = "Introduction" });
        model.Sections.Add(new SectionItem { Number = 2, Title = "Overview" });
        model.Sections.Add(new SectionItem { Number = 3, Title = "Details" });
        model.Sections.Add(new SectionItem { Number = 4, Title = "Conclusion" });
        model.Sections.Add(new SectionItem { Number = 5, Title = "Appendix" });

        // 4. Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 5. Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
