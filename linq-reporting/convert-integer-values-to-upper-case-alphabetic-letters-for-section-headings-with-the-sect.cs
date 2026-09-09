using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Add a heading that will display the section number as an uppercase letter.
        // The model provides a computed property SectionLetter for this purpose.
        builder.Writeln("<<[model.SectionLetter]>>. Section Heading");

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // 2. Load the template document for reporting.
        Document reportDoc = new Document(templatePath);

        // 3. Prepare sample data.
        // The model contains an integer Section and a derived property SectionLetter.
        var model = new ReportModel { Section = 3 }; // Will be displayed as "C"

        // 4. Build the report using Aspose.Words LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // 5. Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);

        Console.WriteLine($"Report generated: {reportPath}");
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Integer section number.
    public int Section { get; set; }

    // Computed property that converts the integer to an uppercase alphabetic letter (A‑Z).
    // Values outside 1‑26 are wrapped around (e.g., 27 -> A).
    public string SectionLetter
    {
        get
        {
            if (Section <= 0)
                return string.Empty;

            // Wrap the number to the range 1‑26.
            int index = ((Section - 1) % 26) + 1;
            char letter = (char)('A' + index - 1);
            return letter.ToString();
        }
    }
}
