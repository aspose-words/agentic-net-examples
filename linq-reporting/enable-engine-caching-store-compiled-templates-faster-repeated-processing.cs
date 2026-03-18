using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class EngineCachingExample
{
    static void Main()
    {
        // Load the template document (create/load rule)
        Document template = new Document("Template.docx");

        // Enable caching of compiled templates.
        // The ReportingEngine uses reflection optimization and caches compiled templates internally.
        // Setting this static property to true ensures the cache is used (default is true).
        ReportingEngine.UseReflectionOptimization = true;

        // Create the reporting engine (create rule)
        ReportingEngine engine = new ReportingEngine();

        // Prepare data source for the template.
        DataSet data = GetData();

        // Build the report (engine uses the cached compiled template on subsequent calls).
        engine.BuildReport(template, data);

        // Save the generated report (save rule)
        template.Save("Report.docx");
    }

    // Example method to obtain a DataSet; replace with actual data retrieval logic.
    static DataSet GetData()
    {
        DataSet ds = new DataSet();

        DataTable table = new DataTable("Employees");
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Position", typeof(string));

        table.Rows.Add("John Doe", "Developer");
        table.Rows.Add("Jane Smith", "Designer");

        ds.Tables.Add(table);
        return ds;
    }
}
