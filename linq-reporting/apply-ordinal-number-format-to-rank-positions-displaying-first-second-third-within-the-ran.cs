using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class RankingItem
{
    public int Position { get; set; }
    public string Name { get; set; } = "";
}

public class ReportModel
{
    public List<RankingItem> Rankings { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create template document
        var templatePath = "template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Ranking Report");
        builder.Writeln("");
        builder.Writeln("<<foreach [item in Rankings]>>");
        builder.Writeln(
            "<<if [item.Position == 1]>>First<</if>>" +
            "<<if [item.Position == 2]>>Second<</if>>" +
            "<<if [item.Position == 3]>>Third<</if>>" +
            "<<if [item.Position > 3]>><<[item.Position]>>th<</if>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Document.Save(templatePath);

        // Load template
        var doc = new Document(templatePath);

        // Prepare data
        var model = new ReportModel();
        model.Rankings.Add(new RankingItem { Position = 1, Name = "Alice" });
        model.Rankings.Add(new RankingItem { Position = 2, Name = "Bob" });
        model.Rankings.Add(new RankingItem { Position = 3, Name = "Charlie" });
        model.Rankings.Add(new RankingItem { Position = 4, Name = "Diana" });
        model.Rankings.Add(new RankingItem { Position = 5, Name = "Ethan" });

        // Build report
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save output
        var outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
