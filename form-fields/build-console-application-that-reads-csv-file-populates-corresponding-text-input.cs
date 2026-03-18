using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Fields; // Added for FormField and FieldType

namespace CsvToFormFields
{
    class Program
    {
        static void Main(string[] args)
        {
            // Expect three arguments: template.docx, data.csv, output.docx
            if (args.Length != 3)
            {
                Console.WriteLine("Usage: CsvToFormFields <templatePath> <csvPath> <outputPath>");
                return;
            }

            string templatePath = args[0];
            string csvPath = args[1];
            string outputPath = args[2];

            // Load the Word template.
            Document doc = new Document(templatePath);

            // Configure CSV loading options.
            // Assume the first row contains column headers.
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
            // You can customize delimiter, comment char etc. if needed:
            // loadOptions.Delimiter = ',';
            // loadOptions.CommentChar = '#';

            // Create a CSV data source.
            CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

            // Build the report. The data source name ("Data") is used in the template tags.
            // Example tag in the template: <<[Data.FirstName]>>
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "Data");

            // OPTIONAL: If the template contains form fields (text input fields) whose
            // names match the CSV column names, copy the merged values into them.
            foreach (FormField field in doc.Range.FormFields)
            {
                // Process only text input form fields.
                if (field.Type == FieldType.FieldFormTextInput)
                {
                    var bookmark = doc.Range.Bookmarks[field.Name];
                    if (bookmark != null)
                    {
                        // Bookmark.Text may be null, guard against it.
                        field.Result = bookmark.Text ?? string.Empty;
                    }
                }
            }

            // Save the populated document.
            doc.Save(outputPath);
        }
    }
}
