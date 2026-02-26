using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Reporting;

namespace AsposeWordsVariableDemo
{
    // Simple data model for the report.
    public class ReportData
    {
        public string Title { get; set; }
        public string Date { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a DOCX template with document variables.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Add two document variables (initial empty values).
            template.Variables.Add("ReportTitle", string.Empty);
            template.Variables.Add("ReportDate", string.Empty);

            // Insert a DOCVARIABLE field for the title.
            FieldDocVariable titleField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
            titleField.VariableName = "ReportTitle";
            titleField.Update();

            builder.Writeln(); // New line between fields.

            // Insert a DOCVARIABLE field for the date.
            FieldDocVariable dateField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
            dateField.VariableName = "ReportDate";
            dateField.Update();

            // Save the template to disk.
            const string templatePath = "ReportTemplate.docx";
            template.Save(templatePath);

            // 2. Load the template (simulating a real scenario where the template is pre‑created).
            Document doc = new Document(templatePath);

            // 3. Prepare a collection of data using LINQ.
            List<ReportData> data = new List<ReportData>
            {
                new ReportData { Title = "Quarterly Sales Report", Date = DateTime.Now.ToString("MMMM dd, yyyy") },
                new ReportData { Title = "Annual Financial Summary", Date = DateTime.Now.AddMonths(-1).ToString("MMMM dd, yyyy") }
            };

            // 4. For each data item, set the document variables, update fields and save a separate report.
            int index = 1;
            foreach (ReportData item in data)
            {
                // Set the values of the document variables.
                doc.Variables["ReportTitle"] = item.Title;
                doc.Variables["ReportDate"] = item.Date;

                // Update all fields so that DOCVARIABLE fields reflect the new values.
                doc.UpdateFields();

                // Optionally, you can also use the ReportingEngine if you want to combine
                // variable usage with LINQ template syntax. Here we demonstrate a simple build:
                ReportingEngine engine = new ReportingEngine();
                // The data source name "data" allows referencing the whole object in the template,
                // but since we already used DOCVARIABLE fields, this call is optional.
                engine.BuildReport(doc, item, "data");

                // Save the generated report.
                string outputPath = $"Report_{index}.docx";
                doc.Save(outputPath);
                index++;
            }
        }
    }
}
