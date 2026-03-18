using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

class CsvRecordFetcher
{
    static void Main()
    {
        // Load a template document that contains a table placeholder for the CSV data.
        // The template should have a MERGEFIELD or a Reporting Engine tag that will be filled with the CSV rows.
        Document doc = new Document("Template.docx");

        // Configure CSV loading options – the first line contains column headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        // Adjust delimiters if needed, e.g., loadOptions.Delimiter = ';';

        // Create a CSV data source from the file.
        CsvDataSource csvSource = new CsvDataSource("Data.csv", loadOptions);

        // Build the report – this will populate the table in the template with the CSV rows.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, csvSource, "CsvData");

        // After building the report, the document now contains a table with the CSV records.
        // Access the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Specify the zero‑based index of the record we want to fetch.
        int recordIndex = 2; // for example, the third record in the CSV.

        // Use the indexer (ElementAt) to retrieve the specific row.
        Row selectedRow = table.Rows[recordIndex];

        // Display the contents of each cell in the selected row.
        Console.WriteLine($"Record #{recordIndex + 1}:");
        for (int i = 0; i < selectedRow.Cells.Count; i++)
        {
            // Convert the cell to plain text.
            string cellText = selectedRow.Cells[i].ToString(SaveFormat.Text).Trim();
            Console.WriteLine($"  Column {i + 1}: {cellText}");
        }

        // Save the document with the populated data (optional).
        doc.Save("CsvReportWithSelectedRecord.docx");
    }
}
