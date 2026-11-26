using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Sections = new List<Section>
            {
                new Section { Number = 1 },
                new Section { Number = 2 },
                new Section { Number = 3 },
                new Section { Number = 4 },
                new Section { Number = 5 }
            }
        };

        // Create a template document programmatically.
        const string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Sections Report");
        builder.Writeln();

        // Begin foreach loop over Sections.
        builder.Writeln("<<foreach [sec in Sections]>>");
        // Use the computed SectionLetter property for heading.
        builder.Writeln("<<[sec.SectionLetter]>>. Section <<[sec.Number]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var templateDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        const string outputPath = "output.docx";
        templateDoc.Save(outputPath);
    }
}

// Wrapper model for the report.
public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

// Section data model.
public class Section
{
    public int Number { get; set; }

    // Convert the integer to an upper‑case alphabetic letter (A, B, C, ...).
    public string SectionLetter => ((char)('A' + (Number - 1))).ToString();
}
