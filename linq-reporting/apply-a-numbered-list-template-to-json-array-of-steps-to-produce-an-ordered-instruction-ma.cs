using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Step
{
    public string Description { get; set; } = string.Empty;
}

public class Model
{
    public List<Step> Steps { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data representing an array of steps.
        const string jsonFile = "steps.json";
        string jsonContent = @"{
            ""Steps"": [
                { ""Description"": ""Preheat the oven to 180°C (350°F)."" },
                { ""Description"": ""Mix flour, sugar, and butter in a bowl."" },
                { ""Description"": ""Add eggs and stir until smooth."" },
                { ""Description"": ""Pour the batter into a greased pan."" },
                { ""Description"": ""Bake for 25 minutes or until golden brown."" }
            ]
        }";
        File.WriteAllText(jsonFile, jsonContent);

        // Deserialize JSON into a strongly‑typed model.
        Model model = JsonConvert.DeserializeObject<Model>(jsonContent) ?? new Model();

        // Create a new blank document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a numbered list style to the paragraphs that will contain the steps.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = numberedList;

        // Insert LINQ Reporting tags.
        // <<restartNum>> must be placed immediately before the <<foreach>> tag in the same numbered paragraph.
        builder.Writeln("<<restartNum>><<foreach [step in model.Steps]>>");
        // Each iteration writes the step description as a separate numbered paragraph.
        builder.Writeln("<<[step.Description]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the template and the strongly‑typed data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated instruction manual.
        doc.Save("InstructionManual.docx");
    }
}
