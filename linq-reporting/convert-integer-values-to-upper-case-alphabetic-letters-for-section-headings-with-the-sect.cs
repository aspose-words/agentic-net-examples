using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // ---------- Create the template ----------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the Sections collection
        builder.Writeln("<<foreach [sec in Sections]>>");

        // Use a custom property that already contains the uppercase letter
        builder.Writeln("<<[sec.SectionLetter]>>. <<[sec.Title]>>");

        // End the foreach loop
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // ---------- Load the template ----------
        Document doc = new Document(templatePath);

        // Prepare sample data
        ReportModel model = new ReportModel
        {
            Sections = new List<SectionItem>
            {
                new SectionItem { Section = 1, Title = "Introduction" },
                new SectionItem { Section = 2, Title = "Details" },
                new SectionItem { Section = 3, Title = "Conclusion" }
            }
        };

        // Build the report using the LINQ Reporting engine
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        doc.Save(reportPath);

        // Inform the user (no interactive input required)
        Console.WriteLine($"Report generated: {reportPath}");
    }
}

// Root data model
public class ReportModel
{
    public List<SectionItem> Sections { get; set; } = new();
}

// Individual section item
public class SectionItem
{
    // Numeric section index
    public int Section { get; set; }

    // Title of the section
    public string Title { get; set; } = string.Empty;

    // Upper‑case alphabetic representation of the Section number (A, B, C, …)
    public string SectionLetter => ((char)('A' + Section - 1)).ToString();
}
