using System;
using System.IO;
using System.Linq;

class CsvRecordFetcher
{
    static void Main()
    {
        const string csvPath = "Data.csv";

        if (!File.Exists(csvPath))
        {
            Console.WriteLine($"CSV file not found: {csvPath}");
            return;
        }

        // Read all lines from the CSV file.
        string[] allLines = File.ReadAllLines(csvPath);

        if (allLines.Length == 0)
        {
            Console.WriteLine("CSV file is empty.");
            return;
        }

        // Assume the first line contains column headers.
        string headerLine = allLines[0];
        string[] headers = headerLine.Split(',');

        // Parse the remaining lines into a collection of string arrays.
        var records = allLines
            .Skip(1) // skip header
            .Select(line => line.Split(','))
            .ToList();

        // Specify the zero‑based index of the record we want to fetch.
        int recordIndex = 2; // for example, the third record in the CSV.

        if (recordIndex < 0 || recordIndex >= records.Count)
        {
            Console.WriteLine($"Record index {recordIndex} is out of range. Total records: {records.Count}");
            return;
        }

        // Use ElementAt to retrieve the specific record.
        string[] selectedRecord = records.ElementAt(recordIndex);

        // Display the contents of each column in the selected record.
        Console.WriteLine($"Record #{recordIndex + 1}:");
        for (int i = 0; i < selectedRecord.Length; i++)
        {
            string columnName = i < headers.Length ? headers[i] : $"Column{i + 1}";
            Console.WriteLine($"  {columnName}: {selectedRecord[i].Trim()}");
        }
    }
}
