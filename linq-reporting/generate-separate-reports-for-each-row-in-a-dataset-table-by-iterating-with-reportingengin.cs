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

        // Create a template document containing LINQ Reporting tags.
        string templatePath = "Template.docx";
        CreateTemplateDocument(templatePath);

        // Iterate over each DataRow and generate a separate report.
        int index = 0;
        foreach (DataRow row in dataSet.Tables["Employees"].Rows)
        {
            // Load a fresh copy of the template for each iteration.
            Document report = new Document(templatePath);

            // Build the report using the current DataRow as the data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(report, row);

            // Save the generated report to a distinct file.
            string outputPath = $"Report_{index}.docx";
            report.Save(outputPath);
            index++;
        }
    }

    // Creates a DataSet with a single DataTable named "Employees".
    private static DataSet CreateSampleDataSet()
    {
        DataTable table = new DataTable("Employees");
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Position", typeof(string));

        table.Rows.Add("Alice Johnson", "Project Manager");
        table.Rows.Add("Bob Smith", "Software Engineer");
        table.Rows.Add("Carol White", "Designer");

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
        // LINQ Reporting tags reference column names directly for DataRow sources.
        builder.Writeln("Name: <<[Name]>>");
        builder.Writeln("Position: <<[Position]>>");

        doc.Save(filePath);
    }
}
