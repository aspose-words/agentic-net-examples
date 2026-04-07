using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Section
{
    public int Number { get; set; }

    // Returns the section number as an uppercase alphabetic letter (A, B, C, ...).
    public string Letter => ((char)('A' + Number - 1)).ToString();

    public string Title { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Sections =
            {
                new Section { Number = 1, Title = "Introduction" },
                new Section { Number = 2, Title = "Methodology" },
                new Section { Number = 3, Title = "Results" },
                new Section { Number = 4, Title = "Conclusion" }
            }
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a heading for each section.
        // The tag <<foreach [sec in Sections]>> iterates over the collection.
        // <<[sec.Letter]>> provides the uppercase alphabetic representation of the section number.
        builder.Writeln("<<foreach [sec in Sections]>>");
        builder.Writeln("<<[sec.Letter]>>. <<[sec.Title]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("Report_Sections.docx");
    }
}
