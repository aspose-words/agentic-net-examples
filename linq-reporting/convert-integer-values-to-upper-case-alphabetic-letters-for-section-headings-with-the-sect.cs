using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create the template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Begin a foreach loop over the Sections collection.
        builder.Writeln("<<foreach [sec in Sections]>>");

        // Write a heading where the integer Section value is converted to an uppercase letter.
        builder.Writeln("<<[sec.Letter]>>. <<[sec.Title]>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template back for report generation.
        var document = new Document(templatePath);

        // Prepare the data model.
        var model = new ReportModel
        {
            Sections = new List<SectionItem>
            {
                new SectionItem { Section = 1, Title = "Introduction" },
                new SectionItem { Section = 2, Title = "Details" },
                new SectionItem { Section = 3, Title = "Conclusion" }
            }
        };

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(document, model, "model");

        // Save the generated report.
        const string reportPath = "Report.docx";
        document.Save(reportPath);
    }
}

// Root data model for the report.
public class ReportModel
{
    // Collection of sections to be iterated over in the template.
    public List<SectionItem> Sections { get; set; } = new();
}

// Represents a single section with an integer identifier and a title.
public class SectionItem
{
    // Integer value that will be converted to an uppercase letter in the report.
    public int Section { get; set; }

    // Title of the section.
    public string Title { get; set; } = string.Empty;

    // Computed property that converts the integer Section to an uppercase alphabetic letter (1 → A, 2 → B, etc.).
    public string Letter => ((char)('A' + Section - 1)).ToString();
}
