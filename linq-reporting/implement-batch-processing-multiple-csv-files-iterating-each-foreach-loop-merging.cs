using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsBatchMailMerge
{
    class Program
    {
        static void Main()
        {
            // Folder that contains the CSV files to be processed.
            string csvFolder = @"C:\InputCsv";
            // Path to the Word template that contains MERGEFIELDs matching the CSV column names.
            string templatePath = @"C:\Template\MailMergeTemplate.docx";
            // Path where the final merged document will be saved.
            string outputPath = @"C:\Output\MergedResult.docx";

            // Load the template once – it will be cloned for each CSV file.
            Document template = new Document(templatePath);

            // Create an empty document that will hold the merged results of all CSV files.
            Document mergedResult = new Document();
            // Remove the default empty section that Aspose.Words adds on construction.
            mergedResult.RemoveAllChildren();

            // Get all CSV files in the specified folder.
            string[] csvFiles = Directory.GetFiles(csvFolder, "*.csv");

            foreach (string csvFile in csvFiles)
            {
                // Build a DataTable from the current CSV file.
                DataTable data = BuildDataTableFromCsv(csvFile, ',');

                // Clone the template so each CSV gets its own independent document.
                Document part = (Document)template.Clone(true);

                // Perform the mail merge using the DataTable.
                part.MailMerge.Execute(data);

                // Append the merged part to the final result document.
                mergedResult.AppendDocument(part, ImportFormatMode.KeepSourceFormatting);
            }

            // Save the combined document.
            mergedResult.Save(outputPath, SaveFormat.Docx);
        }

        /// <summary>
        /// Reads a CSV file and returns a DataTable.
        /// The first line is assumed to contain column headers.
        /// </summary>
        /// <param name="csvPath">Full path to the CSV file.</param>
        /// <param name="delimiter">Character that separates fields (e.g., ',' or ';').</param>
        /// <returns>DataTable populated with the CSV data.</returns>
        private static DataTable BuildDataTableFromCsv(string csvPath, char delimiter)
        {
            DataTable table = new DataTable();

            using (var reader = new StreamReader(csvPath))
            {
                bool isFirstLine = true;
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    if (string.IsNullOrWhiteSpace(line))
                        continue; // Skip empty lines.

                    string[] fields = line.Split(delimiter);

                    if (isFirstLine)
                    {
                        // Create columns using the header row.
                        foreach (string header in fields)
                        {
                            string columnName = header.Trim();
                            if (string.IsNullOrEmpty(columnName))
                                columnName = $"Column{table.Columns.Count + 1}";
                            table.Columns.Add(columnName, typeof(string));
                        }
                        isFirstLine = false;
                    }
                    else
                    {
                        // Add a new row with the field values.
                        DataRow row = table.NewRow();
                        for (int i = 0; i < table.Columns.Count && i < fields.Length; i++)
                        {
                            row[i] = fields[i].Trim();
                        }
                        table.Rows.Add(row);
                    }
                }
            }

            return table;
        }
    }
}
