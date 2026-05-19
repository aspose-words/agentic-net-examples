using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Collection of tags to be displayed.
    public List<string> Tags { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // ---------- Create template ----------
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Show tag list only when the collection contains at least one element.
        builder.Writeln("<<if [model.Tags.Any()]>>");
        builder.Writeln("Tag list:");
        builder.Writeln("<<foreach [tag in model.Tags]>>- <<[tag]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</if>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // ---------- Load template ----------
        var loadedDoc = new Document(templatePath);

        // ---------- Prepare data ----------
        var model = new ReportModel
        {
            Tags = new List<string> { "Tag1", "Tag2", "Tag3" }
        };

        // ---------- Build report ----------
        var engine = new ReportingEngine();
        engine.BuildReport(loadedDoc, model, "model");

        // ---------- Save result ----------
        var outputPath = "Report.docx";
        loadedDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine("Report generated: " + outputPath);
    }
}
