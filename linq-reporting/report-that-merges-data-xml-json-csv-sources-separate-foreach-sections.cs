using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReportGenerationExample
{
    class Program
    {
        static void Main()
        {
            // Path to the template document that contains three separate foreach sections:
            // <<foreach [xml]>><<[Name]>> <<endforeach>>
            // <<foreach [json]>><<[Title]>> <<endforeach>>
            // <<foreach [csv]>><<[Product]>> <<endforeach>>
            string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create data sources for XML, JSON and CSV files.
            // The files must exist and contain data that matches the fields used in the template.
            string xmlPath = @"C:\Data\People.xml";
            string jsonPath = @"C:\Data\Books.json";
            string csvPath = @"C:\Data\Products.csv";

            // XmlDataSource – default loading options.
            XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

            // JsonDataSource – default loading options.
            JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

            // CsvDataSource – specify that the first line contains column headers.
            CsvDataLoadOptions csvLoadOptions = new CsvDataLoadOptions(hasHeaders: true);
            CsvDataSource csvDataSource = new CsvDataSource(csvPath, csvLoadOptions);

            // Prepare arrays of data sources and their corresponding names.
            // The first name can be empty if you only need to reference members,
            // but we give explicit names for clarity.
            object[] dataSources = { xmlDataSource, jsonDataSource, csvDataSource };
            string[] dataSourceNames = { "xml", "json", "csv" };

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSources, dataSourceNames);

            // Save the populated report.
            string outputPath = @"C:\Output\CombinedReport.docx";
            doc.Save(outputPath);
        }
    }
}
