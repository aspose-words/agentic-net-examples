using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "InstructionTemplate.docx");
        string jsonPath = Path.Combine(workDir, "steps.json");
        string outputPath = Path.Combine(workDir, "InstructionManual.docx");

        // ---------- Create the template document ----------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title
        builder.Writeln("Instruction Manual:");

        // Create a numbered list
        List numberedList = templateDoc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = numberedList;

        // Start the numbered paragraph with restartNum and foreach tags
        builder.Writeln("<<restartNum>><<foreach [step in steps]>>");

        // Paragraph that will be repeated for each step
        builder.Writeln("<<[step.Description]>>");

        // End of the foreach block
        builder.Writeln("<</foreach>>");

        // Finish the list
        builder.ListFormat.RemoveNumbers();

        // Save the template to disk (required by lifecycle rule)
        templateDoc.Save(templatePath);

        // ---------- Create sample JSON data ----------
        string jsonContent = @"[
            { ""Description"": ""Preheat the oven to 180°C."" },
            { ""Description"": ""Mix flour, sugar, and eggs together."" },
            { ""Description"": ""Pour the batter into a greased pan."" },
            { ""Description"": ""Bake for 30 minutes."" },
            { ""Description"": ""Let it cool before serving."" }
        ]";
        File.WriteAllText(jsonPath, jsonContent);

        // ---------- Load the template ----------
        Document doc = new Document(templatePath);

        // Load JSON data source
        JsonDataSource dataSource = new JsonDataSource(jsonPath);

        // Build the report using the data source name "steps"
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "steps");

        // Save the final document
        doc.Save(outputPath);
    }
}
