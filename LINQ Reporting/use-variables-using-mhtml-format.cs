using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some custom variables to the document.
        doc.Variables.Add("CompanyName", "Acme Corp");
        doc.Variables.Add("ReportDate", DateTime.Now.ToString("yyyy-MM-dd"));

        // Insert fields that will display the variable values.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Company: ");
        FieldDocVariable companyField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        companyField.VariableName = "CompanyName";
        companyField.Update();

        builder.Writeln("Date: ");
        FieldDocVariable dateField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        dateField.VariableName = "ReportDate";
        dateField.Update();

        // Configure save options for MHTML output.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Use CID URLs for resources to improve compatibility with some mail agents.
            ExportCidUrlsForMhtmlResources = true,
            // Export fonts and images as separate resources.
            ExportFontResources = true,
            ExportImagesAsBase64 = false,
            // Optional: make the output HTML pretty-formatted.
            PrettyFormat = true
        };

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "DocumentWithVariables.mht");

        // Save the document as MHTML using the configured options.
        doc.Save(outputPath, saveOptions);

        // Output the location of the generated file.
        Console.WriteLine($"MHTML document saved to: {outputPath}");
    }
}
