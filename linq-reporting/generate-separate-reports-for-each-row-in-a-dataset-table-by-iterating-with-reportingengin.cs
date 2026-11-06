using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LinqReportingSeparateRows
{
    public static void Main()
    {
        // Create sample DataSet with one DataTable.
        DataSet dataSet = new DataSet();
        DataTable table = new DataTable("People");
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Age", typeof(int));
        table.Rows.Add("Alice", 30);
        table.Rows.Add("Bob", 45);
        table.Rows.Add("Charlie", 28);
        dataSet.Tables.Add(table);

        // Create a template document with LINQ Reporting tags.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report for <<[row.Name]>>");
        builder.Writeln("Age: <<[row.Age]>>");
        templateDoc.Save(templatePath);

        // Ensure the output directory exists.
        string outputDir = "Reports";
        Directory.CreateDirectory(outputDir);

        // Iterate over each DataRow and generate a separate report.
        foreach (DataRow dataRow in dataSet.Tables["People"].Rows)
        {
            // Load a fresh copy of the template for each report.
            Document reportDoc = new Document(templatePath);

            // Build the report using the current DataRow as the root object.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, dataRow, "row");

            // Save the generated report with a unique file name.
            string name = dataRow["Name"]?.ToString() ?? "Unnamed";
            string reportPath = Path.Combine(outputDir, $"Report_{name}.docx");
            reportDoc.Save(reportPath);
        }
    }
}
