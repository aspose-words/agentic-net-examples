using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}

public class Model
{
    // Collection that will be iterated in the template.
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // ---------- Create the template document ----------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Opening tag for a foreach loop.
        builder.Writeln("<<foreach [item in Items]>>");
        // Content inside the loop.
        builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
        // Closing tag for the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template (demonstrates the save lifecycle rule).
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // ---------- Load the template ----------
        Document doc = new Document(templatePath);

        // ---------- Prepare sample data ----------
        Model model = new Model();
        model.Items.Add(new Item { Index = 1, Name = "Apple" });
        model.Items.Add(new Item { Index = 2, Name = "Banana" });
        model.Items.Add(new Item { Index = 3, Name = "Cherry" });

        // ---------- Build the report ----------
        ReportingEngine engine = new ReportingEngine();
        // BuildReport returns a bool indicating success when InlineErrorMessages option is used.
        // Here we just use the default options; the return value will be true if parsing succeeded.
        bool success = engine.BuildReport(doc, model, "model");

        // ---------- Save the generated report ----------
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
    }
}
