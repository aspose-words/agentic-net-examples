using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingExample
{
    // Simple data source class whose members will be referenced from the DOCX template.
    public class ReportData
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public decimal Salary { get; set; }

        public ReportData(string name, int age, decimal salary)
        {
            Name = name;
            Age = age;
            Salary = salary;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains LINQ Reporting tags, e.g. <<[data.Name]>>.
            const string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Load the template document using the Document constructor (lifecycle rule).
            Document doc = new Document(templatePath);

            // Create an instance of the data source.
            ReportData data = new ReportData("John Doe", 42, 12345.67m);

            // Populate the template with the data using ReportingEngine (feature rule).
            ReportingEngine engine = new ReportingEngine();
            // The second overload allows referencing the data source object itself via the name "data".
            engine.BuildReport(doc, data, "data");

            // Configure HtmlFixedSaveOptions (required for saving to HTML Fixed format).
            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions
            {
                // Ensure the correct format is set (mandatory for the Save method overload).
                SaveFormat = SaveFormat.HtmlFixed,

                // Optional: embed images as Base64 to keep a single HTML file.
                ExportEmbeddedImages = true,

                // Optional: embed CSS to simplify the output folder structure.
                ExportEmbeddedCss = true,

                // Optional: set a pretty format for readability.
                PrettyFormat = true
            };

            // Save the populated document as HTML Fixed using the Save method with options (lifecycle rule).
            const string outputPath = @"C:\Output\Report.html";
            doc.Save(outputPath, saveOptions);
        }
    }
}
