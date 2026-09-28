using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel
        {
            Groups = new List<Group>
            {
                new Group { Name = "Group A", ColumnCount = 3 },
                new Group { Name = "Group B", ColumnCount = 2 },
                new Group { Name = "Group C", ColumnCount = 5 }
            }
        };

        // Determine the maximum number of columns needed.
        int maxColumns = model.Groups.Max(g => g.ColumnCount);

        // Create the template document.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin foreach over Groups.
        builder.Writeln("<<foreach [g in Groups]>>");

        // Start a table for each group.
        Table table = builder.StartTable();

        for (int col = 0; col < maxColumns; col++)
        {
            builder.InsertCell();

            if (col == 0)
            {
                // First cell always contains the merge tag and the group name.
                builder.Writeln("<<cellMerge>><<[g.Name]>>");
            }
            else
            {
                // Subsequent cells are added only if the group's ColumnCount exceeds the current index.
                builder.Writeln($"<<if [g.ColumnCount > {col}]>><<cellMerge>><<[g.Name]>> <</if>>");
            }
        }

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;

        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        var outputPath = Path.Combine("output", "Report.docx");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
        reportDoc.Save(outputPath);

        // Indicate completion (no interactive input).
        Console.WriteLine(success ? "Report generated successfully." : "Report generation failed.");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Group> Groups { get; set; } = new();
}

public class Group
{
    public string Name { get; set; } = "";
    public int ColumnCount { get; set; }
}
