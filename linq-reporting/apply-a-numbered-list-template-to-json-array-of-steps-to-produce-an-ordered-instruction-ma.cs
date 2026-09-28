using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data representing an array of steps
        string json = @"{
            ""Steps"": [
                { ""Description"": ""Preheat the oven to 180°C."" },
                { ""Description"": ""Mix flour and sugar in a bowl."" },
                { ""Description"": ""Add eggs and stir until smooth."" },
                { ""Description"": ""Pour batter into a greased pan."" },
                { ""Description"": ""Bake for 30 minutes."" }
            ]
        }";

        // Deserialize JSON into the data model
        InstructionManual manual = JsonConvert.DeserializeObject<InstructionManual>(json)!;

        // Create the LINQ Reporting template programmatically
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title
        builder.Writeln("Instruction Manual");
        builder.Writeln();

        // Begin foreach over the Steps collection
        builder.Writeln("<<foreach [step in Steps]>>");

        // Set the paragraph style to a numbered list
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.ListNumber;

        // Restart numbering for each step and insert the step description
        builder.Writeln("<<restartNum>><<[step.Description]>>");

        // End foreach block
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // Load the template (ensures BuildReport is called after loading)
        Document doc = new Document(templatePath);

        // Build the report using the data model
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, manual, "manual");

        // Save the generated instruction manual
        string outputPath = "InstructionManual.docx";
        doc.Save(outputPath);
    }
}

// Root data model for the report
public class InstructionManual
{
    public List<Step> Steps { get; set; } = new();
}

// Model representing a single step
public class Step
{
    public string Description { get; set; } = "";
}
