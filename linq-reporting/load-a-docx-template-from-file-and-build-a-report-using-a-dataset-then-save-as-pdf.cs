using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
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

        // Begin a foreach loop over the DataTable named "People".
        builder.Writeln("<<foreach [person in People]>>");
        // Output each person's name and age.
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age:  <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare a DataSet with a DataTable that matches the template.
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
        // 3. Load the template and build the report using ReportingEngine.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        ReportingEngine engine = new ReportingEngine();
        // BuildReport overload that takes the data source without a name.
        engine.BuildReport(reportDoc, dataSet);

        // -----------------------------------------------------------------
        // 4. Save the generated report as PDF.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPdfPath, SaveFormat.Pdf);
    }
}
