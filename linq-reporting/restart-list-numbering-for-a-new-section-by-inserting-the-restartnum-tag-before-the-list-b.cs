using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class ReportModel
{
    public List<string> Items { get; set; } = new();
    public List<string> Items2 { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // -------------------------------------------------
        // Create the LINQ Reporting template programmatically
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // First section with a numbered list
        builder.Writeln("First Section:");
        builder.ListFormat.ApplyNumberDefault(); // start numbered list
        builder.Writeln("<<foreach [item in Items]>><<[item]>> <</foreach>>");
        builder.ListFormat.RemoveNumbers(); // end list

        // Section break to start a new section
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section with a numbered list that restarts numbering
        builder.Writeln("Second Section:");
        builder.ListFormat.ApplyNumberDefault(); // start numbered list
        // <<restartNum>> placed before the foreach tag restarts numbering for this list
        builder.Writeln("<<restartNum>><<foreach [item in Items2]>><<[item]>> <</foreach>>");
        builder.ListFormat.RemoveNumbers(); // end list

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Prepare sample data for the report
        // -------------------------------------------------
        ReportModel model = new ReportModel
        {
            Items = new List<string> { "Apple", "Banana", "Cherry" },
            Items2 = new List<string> { "Dog", "Elephant", "Frog" }
        };

        // -------------------------------------------------
        // Load the template and build the report
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        reportDoc.Save(outputPath);
    }
}
