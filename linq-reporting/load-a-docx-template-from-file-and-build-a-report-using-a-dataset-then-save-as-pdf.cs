using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LinqReportingExample
{
    public static void Main()
    {
        // Paths for the template and the final PDF.
        const string templatePath = "Template.docx";
        const string outputPdfPath = "Report.pdf";

        // -----------------------------------------------------------------
        // 1. Create a DOCX template with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a title.
        builder.Writeln("People Report");
        builder.Writeln();

        // Begin a foreach loop over the DataSet's table "People".
        // The root object name is "ds" (for DataSet) and the table name is "People".
        builder.Writeln("<<foreach [person in ds.People]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age:  <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare a DataSet with sample data.
        // -----------------------------------------------------------------
        DataSet dataSet = new DataSet();

        DataTable peopleTable = new DataTable("People");
        peopleTable.Columns.Add("Name", typeof(string));
        peopleTable.Columns.Add("Age", typeof(int));

        peopleTable.Rows.Add("Alice", 30);
        peopleTable.Rows.Add("Bob",   45);
        peopleTable.Rows.Add("Carol", 27);

        dataSet.Tables.Add(peopleTable);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this simple example.
        engine.Options = ReportBuildOptions.None;

        // Build the report using the DataSet as the data source.
        // The root name "ds" must match the name used in the template tags.
        engine.BuildReport(reportDoc, dataSet, "ds");

        // -----------------------------------------------------------------
        // 4. Save the generated report as PDF.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPdfPath, SaveFormat.Pdf);
    }
}
