// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using System.Linq;
using Aspose.Words;

namespace AsposeWordsReporting
{
    public class ReportPrinter
    {
        /// <summary>
        /// Loads a DOCX template, replaces placeholders using LINQ data, and prints the document.
        /// </summary>
        /// <param name="templatePath">Full path to the DOCX template.</param>
        /// <param name="printerName">Name of the printer to use. If null or empty, the default printer is used.</param>
        public static void PrintReport(string templatePath, string printerName = null)
        {
            // Load the existing DOCX document.
            Document doc = new Document(templatePath);

            // Example data source – could be any IEnumerable<T>.
            var customers = new[]
            {
                new { FullName = "John Doe", Address = "123 Main St", City = "New York" },
                new { FullName = "Jane Smith", Address = "456 Oak Ave", City = "Los Angeles" }
            };

            // For demonstration, we will use the first record.
            var record = customers.FirstOrDefault();
            if (record != null)
            {
                // Replace placeholders in the document with actual values.
                // Placeholders in the template should be like {{FullName}}, {{Address}}, {{City}}.
                doc.Range.Replace("{{FullName}}", record.FullName);
                doc.Range.Replace("{{Address}}", record.Address);
                doc.Range.Replace("{{City}}", record.City);
            }

            // If a specific printer is supplied, configure PrinterSettings.
            if (!string.IsNullOrEmpty(printerName))
            {
                PrinterSettings settings = new PrinterSettings
                {
                    PrinterName = printerName,
                    // Example: print all pages; modify as needed.
                    PrintRange = PrintRange.AllPages
                };

                // Print using the specified printer settings.
                doc.Print(settings);
            }
            else
            {
                // Print using the default printer.
                doc.Print();
            }
        }
    }
}
