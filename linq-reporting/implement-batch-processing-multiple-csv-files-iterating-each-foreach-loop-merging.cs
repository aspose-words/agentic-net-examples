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
            // Base directory for all temporary files.
            string baseDir = AppContext.BaseDirectory;

            // Prepare folders.
            string csvFolder = Path.Combine(baseDir, "InputCsv");
            string templateFolder = Path.Combine(baseDir, "Template");
            string outputFolder = Path.Combine(baseDir, "Output");

            Directory.CreateDirectory(csvFolder);
            Directory.CreateDirectory(templateFolder);
            Directory.CreateDirectory(outputFolder);

            // Paths.
            string templatePath = Path.Combine(templateFolder, "MailMergeTemplate.docx");
            string outputPath = Path.Combine(outputFolder, "MergedResult.docx");

            // Ensure a simple template exists.
            if (!File.Exists(templatePath))
            {
                var templateDoc = new Document();
                var builder = new DocumentBuilder(templateDoc);
                builder.Font.Size = 14;
                builder.Writeln("Mail Merge Result");
                builder.InsertParagraph();
                builder.InsertField("MERGEFIELD Name");
                builder.InsertParagraph();
                builder.InsertField("MERGEFIELD Age");
                templateDoc.Save(templatePath);
            }

            // Ensure at least one CSV file exists.
            string[] existingCsv = Directory.GetFiles(csvFolder, "*.csv");
            if (existingCsv.Length == 0)
            {
                string sampleCsvPath = Path.Combine(csvFolder, "Sample1.csv");
                File.WriteAllLines(sampleCsvPath, new[]
                {
                    "Name,Age",
                    "Alice,30",
                    "Bob,25"
                });
            }

            // Load the template once – it will be cloned for each CSV file.
            Document template = new Document(templatePath);

            // Create an empty document that will hold the merged results of all CSV files.
            Document mergedResult = new Document();
            mergedResult.RemoveAllChildren(); // Remove the default empty section.

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

            Console.WriteLine($"Merged document saved to: {outputPath}");
        }

        /// <summary>
        /// Reads a CSV file and returns a DataTable.
        /// The first line is assumed to contain column headers.
        /// </summary>
        private static DataTable BuildDataTableFromCsv(string csvPath, char delimiter)
        {
            var table = new DataTable();

            using (var reader = new StreamReader(csvPath))
            {
                bool isFirstLine = true;
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    string[] fields = line.Split(delimiter);

                    if (isFirstLine)
                    {
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
                        var row = table.NewRow();
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
