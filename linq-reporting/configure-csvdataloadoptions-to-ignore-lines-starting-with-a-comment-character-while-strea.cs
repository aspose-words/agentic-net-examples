using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class CsvCommentExample
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for temporary files.
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("People Report");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");
        template.Save(templatePath);

        // Prepare CSV data with comment lines.
        string csvContent = @"Name,Age
# This line is a comment and should be ignored
Alice,30
Bob,25
# Another comment
Charlie,35";

        // Write CSV data to a memory stream.
        using (MemoryStream csvStream = new MemoryStream(Encoding.UTF8.GetBytes(csvContent)))
        {
            // Configure CSV loading options to recognize headers and comment lines.
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
            {
                Delimiter = ',',
                CommentChar = '#'
            };

            // Create a CSV data source from the stream with the specified options.
            CsvDataSource dataSource = new CsvDataSource(csvStream, loadOptions);

            // Load the template document (demonstrating the load step).
            Document doc = new Document(templatePath);

            // Build the report using the data source. The root name in the template is "persons".
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "persons");

            // Save the generated report.
            doc.Save(outputPath);
        }

        // Clean up temporary template file.
        if (File.Exists(templatePath))
            File.Delete(templatePath);
    }
}
