using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // ---------- 1. Prepare JSON data ----------
        // Simple JSON array of instruction steps.
        string json = "[\"Preheat oven\",\"Mix ingredients\",\"Bake for 30 minutes\"]";
        using MemoryStream jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json));
        // Create a JsonDataSource from the stream.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonStream);

        // ---------- 2. Create the template document ----------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a numbered list style to the paragraph that will contain the steps.
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

        // The <<restartNum>> tag must be placed immediately before the <<foreach>> tag
        // inside the same numbered paragraph to start numbering from 1.
        // Each iteration writes the current step text.
        builder.Writeln("<<restartNum>><<foreach [step in steps]>><<[step]>> <</foreach>>");

        // ---------- 3. Build the report ----------
        ReportingEngine engine = new ReportingEngine();
        // The data source name ("steps") must match the name used in the template tags.
        engine.BuildReport(doc, jsonDataSource, "steps");

        // ---------- 4. Save the resulting instruction manual ----------
        string outputPath = "InstructionManual.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Instruction manual generated: {Path.GetFullPath(outputPath)}");
    }
}
