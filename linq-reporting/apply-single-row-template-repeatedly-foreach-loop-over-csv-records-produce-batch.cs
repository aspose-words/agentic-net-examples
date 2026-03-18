using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class BatchDocumentGenerator
{
    static void Main()
    {
        // Path to the Word template that contains merge fields matching CSV column names.
        const string templatePath = @"C:\Templates\SingleRowTemplate.docx";

        // Path to the CSV file. The first line must contain column headers.
        const string csvPath = @"C:\Data\Records.csv";

        // Load the template document once. It will be cloned for each CSV record.
        Document template = new Document(templatePath);

        // Read the CSV file.
        using (var reader = new StreamReader(csvPath))
        {
            // Parse the header line to obtain field names.
            string headerLine = reader.ReadLine();
            if (headerLine == null) throw new InvalidOperationException("CSV file is empty.");

            // Assuming comma as delimiter; adjust if needed.
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
                string outputPath = $@"C:\Output\Document_{recordIndex:D4}.docx";
                doc.Save(outputPath);

                recordIndex++;
            }
        }
    }
}
