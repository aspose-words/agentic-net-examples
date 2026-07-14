using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class InstructionManualGenerator
{
    // Model that matches the JSON structure.
    public class InstructionData
    {
        public List<string> Steps { get; set; } = new();
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare sample JSON data file.
        // -----------------------------------------------------------------
        string jsonPath = "steps.json";
        string jsonContent = @"{
  ""Steps"": [
    ""Preheat the oven to 180°C."",
    ""Mix flour and sugar in a bowl."",
    ""Add eggs and whisk thoroughly."",
    ""Pour batter into a pan and bake for 30 minutes."",
    ""Let it cool before serving.""
  ]
}";
        File.WriteAllText(jsonPath, jsonContent);

        // -----------------------------------------------------------------
        // 2. Load the JSON into a strongly‑typed model.
        // -----------------------------------------------------------------
        InstructionData model = JsonConvert.DeserializeObject<InstructionData>(jsonContent)!;

        // -----------------------------------------------------------------
        // 3. Create a blank Word document that will serve as the template.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a numbered list style to the paragraph that will contain the loop.
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Insert the LINQ Reporting tags.
        // <<restartNum>> ensures numbering starts at 1 for this list.
        // The foreach iterates over the Steps collection from the model.
        builder.Writeln("<<restartNum>><<foreach [step in Steps]>> <<[step]>> <</foreach>>");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // -----------------------------------------------------------------
        // 4. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // Pass the model as the data source and give it a name ("model").
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated instruction manual.
        // -----------------------------------------------------------------
        doc.Save("InstructionManual.docx");
    }
}
