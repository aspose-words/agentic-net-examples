using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words operations).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample JSON data representing the steps of an instruction manual.
        // -----------------------------------------------------------------
        var steps = new List<Step>
        {
            new Step { Description = "Preheat the oven to 180°C." },
            new Step { Description = "Mix flour, sugar and butter in a bowl." },
            new Step { Description = "Add eggs and whisk until smooth." },
            new Step { Description = "Pour the batter into a greased pan." },
            new Step { Description = "Bake for 25 minutes or until golden brown." }
        };

        string jsonPath = "steps.json";
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(steps, Formatting.Indented));

        // -----------------------------------------------------------------
        // 2. Build the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = "Template.docx";

        // Create a blank document and a builder.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Apply a numbered list style to the paragraph that will contain the loop.
        builder.ListFormat.List = templateDoc.Lists.Add(ListTemplate.NumberDefault);

        // Insert the LINQ Reporting tags.
        // <<restartNum>> ensures numbering starts at 1 for this list.
        // The foreach iterates over the JSON array named "steps".
        // Inside the loop we output the Description property of each step.
        builder.Writeln("<<restartNum>><<foreach [step in steps]>><<[step.Description]>> <</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and generate the report using the JSON data source.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Create a JsonDataSource that reads the JSON file created earlier.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Initialize the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report. The root data source name must match the name used in the template ("steps").
        engine.BuildReport(reportDoc, jsonDataSource, "steps");

        // -----------------------------------------------------------------
        // 4. Save the final instruction manual.
        // -----------------------------------------------------------------
        string outputPath = "InstructionManual.docx";
        reportDoc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model for the JSON array (used only for serialization above).
// ---------------------------------------------------------------------
public class Step
{
    public string Description { get; set; } = string.Empty;
}
