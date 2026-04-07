using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare a simple DataSet with one DataTable named "Items".
        DataSet dataSet = CreateSampleDataSet();

        // Create a template document that contains LINQ Reporting tags.
        string templatePath = "ReportTemplate.docx";
        CreateTemplateDocument(templatePath);

        // Load the template.
        Document templateDoc = new Document(templatePath);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The root object name in the template is "ds".
        engine.BuildReport(templateDoc, dataSet, "ds");

        // Set custom document properties based on values from the DataSet.
        int totalRows = dataSet.Tables["Items"]?.Rows.Count ?? 0;
        string firstTitle = totalRows > 0
            ? dataSet.Tables["Items"]!.Rows[0]["Title"]?.ToString() ?? string.Empty
            : string.Empty;

        templateDoc.CustomDocumentProperties.Add("TotalRows", totalRows);
        templateDoc.CustomDocumentProperties.Add("FirstTitle", firstTitle);
        templateDoc.CustomDocumentProperties.Add("ReportGenerated", DateTime.Now);

        // Save the final report.
        string outputPath = "GeneratedReport.docx";
        templateDoc.Save(outputPath);
    }

    // Creates a DataSet with a single DataTable named "Items".
    private static DataSet CreateSampleDataSet()
    {
        DataTable table = new DataTable("Items");
        table.Columns.Add("Title", typeof(string));
        table.Columns.Add("Value", typeof(string));

        table.Rows.Add("Item A", "123");
        table.Rows.Add("Item B", "456");
        table.Rows.Add("Item C", "789");

        DataSet ds = new DataSet();
        ds.Tables.Add(table);
        return ds;
    }

    // Generates a Word template containing LINQ Reporting tags.
    private static void CreateTemplateDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Simple static title for the report.
        builder.Writeln("Report");
        builder.Writeln();

        // Begin the foreach loop over the Items table.
        builder.Writeln("<<foreach [item in ds.Items]>>");

        // Create the table inside the foreach block.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Title");
        builder.InsertCell();
        builder.Writeln("Value");
        builder.EndRow();

        // Data row for each item.
        builder.InsertCell();
        builder.Writeln("<<[item.Title]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Value]>>");
        builder.EndRow();

        // End the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }
}
