using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data in a DataSet with one DataTable.
        DataSet dataSet = CreateSampleDataSet();

        // Create a LINQ Reporting template document programmatically.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "EmployeeTemplate.docx");
        CreateTemplateDocument(templatePath);

        // Iterate over each DataRow and generate an individual report.
        int reportIndex = 1;
        foreach (DataRow row in dataSet.Tables["Employees"].Rows)
        {
            // Load a fresh copy of the template for each iteration.
            Document report = new Document(templatePath);

            // Build the report using the current DataRow as the data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(report, row);

            // Save the generated report.
            string reportPath = Path.Combine(outputDir, $"EmployeeReport_{reportIndex}.docx");
            report.Save(reportPath);
            reportIndex++;
        }
    }

    // Creates a DataSet containing a single DataTable named "Employees".
    private static DataSet CreateSampleDataSet()
    {
        DataTable table = new DataTable("Employees");
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Position", typeof(string));

        table.Rows.Add(1, "Alice Johnson", "Software Engineer");
        table.Rows.Add(2, "Bob Smith", "Project Manager");
        table.Rows.Add(3, "Carol Davis", "Quality Analyst");

        DataSet ds = new DataSet();
        ds.Tables.Add(table);
        return ds;
    }

    // Generates a simple Word template with LINQ Reporting tags.
    private static void CreateTemplateDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Employee Report");
        builder.Writeln("----------------");
        builder.Writeln("Id: <<[Id]>>");
        builder.Writeln("Name: <<[Name]>>");
        builder.Writeln("Position: <<[Position]>>");

        doc.Save(filePath);
    }
}
