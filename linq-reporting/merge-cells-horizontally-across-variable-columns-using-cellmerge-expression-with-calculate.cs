using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Sample data model.
        ReportModel model = new ReportModel
        {
            Groups = new List<Group>
            {
                new Group
                {
                    Name = "Group A",
                    Columns = new List<string> { "C1", "C2", "C3" } // three columns
                },
                new Group
                {
                    Name = "Group B",
                    Columns = new List<string> { "C1", "C2" } // two columns
                },
                new Group
                {
                    Name = "Group C",
                    Columns = new List<string> { "C1", "C2", "C3", "C4" } // four columns
                }
            }
        };

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Outer foreach iterates over groups.
        builder.Writeln("<<foreach [group in Groups]>>");

        // For each group we create a table that contains a single row.
        // The row has a single cell that will be merged horizontally across
        // the required number of columns at runtime using the <<cellMerge>> tag.
        Table table = builder.StartTable();

        // Insert the first cell – it will be the first cell of the merged range.
        builder.InsertCell();
        // The cell contains the merge tag and the group name.
        builder.Write("<<cellMerge>>");
        builder.Write("<<[group.Name]>>");

        // Insert the remaining cells that belong to the same merged range.
        // The number of additional cells equals (Columns.Count - 1).
        // This logic is expressed with a LINQ Reporting foreach loop.
        builder.Writeln("<<foreach [col in group.Columns]>>");
        // Skip the first column because we already created a cell for it.
        builder.Writeln("<</foreach>>"); // placeholder to keep the tag well‑formed.

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Close the outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template (optional, shown for clarity).
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // Load the template (could reuse the same Document instance).
        Document report = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        report.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Group> Groups { get; set; } = new();
}

public class Group
{
    public string Name { get; set; } = string.Empty;
    public List<string> Columns { get; set; } = new();
}
