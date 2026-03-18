using System;
using Aspose.Words;
using Aspose.Words.Reporting;

// Load the Word template that contains a pre‑formatted table with a list‑in‑table tag.
// Example tag in the template: <<foreach [persons]>><<[Name]>><</foreach>>
string templatePath = @"C:\Templates\ListInTableTemplate.docx";
Document doc = new Document(templatePath);

// Configure CSV loading options – the first line contains column headers.
CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(hasHeaders: true)
{
    Delimiter = ',',          // Adjust if your CSV uses a different delimiter.
    CommentChar = '#',        // Optional: character that starts a comment line.
    QuoteChar = '"'          // Optional: character used for quoting fields.
};

// Create a CSV data source from the file.
string csvPath = @"C:\Data\People.csv";
CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

// Build the report – the data source name ("persons") must match the tag used in the template.
ReportingEngine engine = new ReportingEngine();
engine.BuildReport(doc, csvData, "persons");

// Save the populated document.
string outputPath = @"C:\Output\ListInTableResult.docx";
doc.Save(outputPath);
