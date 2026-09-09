using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Prepare sample data in a DataTable.
        DataTable employees = new DataTable("Employees");
        employees.Columns.Add("Name", typeof(string));
        employees.Columns.Add("Age", typeof(int));
        employees.Rows.Add("Alice", 30);
        employees.Rows.Add("Bob", 45);
        employees.Rows.Add("Charlie", 28);

        // -----------------------------------------------------------------
        // Create a template document programmatically.
        // The template contains LINQ Reporting tags inside content controls.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Employees Report");
        builder.Writeln(); // blank line

        // Start a foreach block that iterates over the DataTable rows.
        builder.Writeln("<<foreach [emp in Employees]>>");

        // Content control for the employee name.
        StructuredDocumentTag nameTag = new StructuredDocumentTag(template, SdtType.PlainText, MarkupLevel.Inline);
        builder.InsertNode(nameTag);
        builder.MoveTo(nameTag);
        builder.Write("<<[emp.Name]>>");

        // Separator.
        builder.Write(" - ");

        // Content control for the employee age.
        StructuredDocumentTag ageTag = new StructuredDocumentTag(template, SdtType.PlainText, MarkupLevel.Inline);
        builder.InsertNode(ageTag);
        builder.MoveTo(ageTag);
        builder.Write("<<[emp.Age]>>");

        // End of the line for each employee.
        builder.Writeln();

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before building the report).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // BuildReport with the DataTable as the data source.
        // The third argument ("Employees") matches the root name used in the tags.
        bool success = engine.BuildReport(report, employees, "Employees");

        // Save the generated report.
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}
