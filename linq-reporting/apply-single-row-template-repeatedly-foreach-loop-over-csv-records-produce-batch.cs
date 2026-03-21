using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class BatchDocumentGenerator
{
    static void Main()
    {
        // Create a temporary working folder.
        string workDir = Path.Combine(Path.GetTempPath(), "LinqReportingDemo");
        Directory.CreateDirectory(workDir);

        // Paths for the template, CSV source and output folder.
        string templatePath = Path.Combine(workDir, "SingleRowTemplate.docx");
        string csvPath = Path.Combine(workDir, "Records.csv");
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Prepare a simple Word template containing merge fields.
        // -----------------------------------------------------------------
        if (!File.Exists(templatePath))
        {
            Document templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Report for <<Name>> (Age: <<Age>>)");
            templateDoc.Save(templatePath);
        }

        // -----------------------------------------------------------------
        // Prepare a small CSV file with header and a few records.
        // -----------------------------------------------------------------
        if (!File.Exists(csvPath))
        {
            File.WriteAllLines(csvPath, new[]
            {
                "Name,Age",
                "Alice,30",
                "Bob,25",
                "Charlie,35"
            });
        }

        // Load the template once; it will be cloned for each CSV record.
        Document template = new Document(templatePath);

        // Read the CSV file.
        using (var reader = new StreamReader(csvPath))
        {
            // Parse the header line to obtain field names.
            string headerLine = reader.ReadLine();
            if (headerLine == null)
                throw new InvalidOperationException("CSV file is empty.");

            string[] fieldNames = headerLine.Split(',');

            int recordIndex = 0;
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                // Split the current line into field values.
                string[] fieldValues = line.Split(',');

                // Clone the template for the current record.
                Document doc = (Document)template.Clone(true);

                // Perform a mail merge for a single record using the field names and values.
                doc.MailMerge.Execute(fieldNames, fieldValues);

                // Save the generated document. Each file gets a unique index.
                string outputPath = Path.Combine(outputDir, $"Document_{recordIndex:D4}.docx");
                doc.Save(outputPath);

                Console.WriteLine($"Generated: {outputPath}");
                recordIndex++;
            }
        }
    }
}
