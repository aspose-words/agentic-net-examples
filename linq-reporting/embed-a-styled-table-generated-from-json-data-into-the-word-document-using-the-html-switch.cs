using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for legacy encodings (required by Aspose.Words on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working folder.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(workDir);

        // File paths.
        string templatePath = Path.Combine(workDir, "template.docx");
        string jsonPath = Path.Combine(workDir, "data.json");
        string resultPath = Path.Combine(workDir, "result.docx");

        // -----------------------------------------------------------------
        // 1. Create sample JSON data containing an HTML table.
        // -----------------------------------------------------------------
        string htmlTable = @"
<table style='border-collapse:collapse; width:50%;'>
    <tr>
        <th style='border:1px solid black; background:#D3D3D3; padding:5px;'>ID</th>
        <th style='border:1px solid black; background:#D3D3D3; padding:5px;'>Name</th>
    </tr>
    <tr>
        <td style='border:1px solid black; padding:5px;'>1</td>
        <td style='border:1px solid black; padding:5px;'>Alice</td>
    </tr>
    <tr>
        <td style='border:1px solid black; padding:5px;'>2</td>
        <td style='border:1px solid black; padding:5px;'>Bob</td>
    </tr>
</table>";

        var jsonObject = new { HtmlTable = htmlTable };
        string jsonContent = JsonConvert.SerializeObject(jsonObject, Formatting.Indented);
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build a template document that contains the LINQ Reporting tag.
        //    The tag uses the html switch to insert the HTML table.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report generated from JSON data:");
        // The tag <<[model.HtmlTable] -html>> tells the engine to evaluate the expression
        // and treat the result as HTML that will be inserted into the document.
        builder.Writeln("<<[model.HtmlTable] -html>>");

        // Save the template before building the report (required by the rules).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and the JSON data source.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // -----------------------------------------------------------------
        // 4. Build the report using ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // Optional: remove empty paragraphs that may appear after processing.
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // The root object name in the template is "model", matching the third argument.
        engine.BuildReport(loadedTemplate, jsonDataSource, "model");

        // -----------------------------------------------------------------
        // 5. Save the final document.
        // -----------------------------------------------------------------
        loadedTemplate.Save(resultPath);
    }
}
