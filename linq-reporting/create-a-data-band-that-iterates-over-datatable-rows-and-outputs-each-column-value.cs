using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data in a DataTable.
        DataTable dt = new DataTable("People");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Age", typeof(int));
        dt.Rows.Add("Alice", 30);
        dt.Rows.Add("Bob", 25);
        dt.Rows.Add("Charlie", 35);

        // Create a blank Word document and insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Data band that iterates over the rows of the DataTable.
        builder.Writeln("<<foreach [row in dt]>>");
        builder.Writeln("Name: <<[row.Name]>>, Age: <<[row.Age]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the DataTable as the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dt, "dt");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
