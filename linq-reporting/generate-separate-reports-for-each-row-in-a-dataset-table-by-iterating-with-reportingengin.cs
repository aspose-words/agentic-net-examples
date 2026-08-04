using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data in a DataSet with one DataTable.
        DataSet dataSet = new DataSet();
        DataTable table = new DataTable("Customers");
        table.Columns.Add("CustomerName", typeof(string));
        table.Columns.Add("Address", typeof(string));
        table.Rows.Add("Thomas Hardy", "120 Hanover Sq., London");
        table.Rows.Add("Paolo Accorti", "Via Monte Bianco 34, Torino");
        dataSet.Tables.Add(table);

        // Create a LINQ Reporting template programmatically.
        const string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Iterate over each DataRow and generate a separate report.
        int index = 1;
        foreach (DataRow row in table.Rows)
        {
            // Load a fresh copy of the template for each iteration.
            Document doc = new Document(templatePath);

            // Build the report using the current DataRow as the data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, row);

            // Save the generated report with a distinct file name.
            string outputPath = $"Report_{index}.docx";
            doc.Save(outputPath);
            index++;
        }
    }

    // Creates a simple Word template containing LINQ Reporting tags.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a title.
        builder.Writeln("Customer Report");
        builder.Writeln("----------------");

        // Insert fields that will be replaced by the data source.
        builder.Writeln("Name: <<[CustomerName]>>");
        builder.Writeln("Address: <<[Address]>>");

        // Save the template to disk.
        doc.Save(filePath);
    }
}
