using System;
using System.IO;
using Aspose.Words;

namespace RevisionExportUtility
{
    /// <summary>
    /// Exports revision metadata from a Word document to a CSV file.
    /// </summary>
    public static class RevisionExporter
    {
        /// <summary>
        /// Loads the document, extracts revision information and writes it to a CSV file.
        /// </summary>
        /// <param name="inputDocPath">Full path to the source .doc/.docx file.</param>
        /// <param name="outputCsvPath">Full path where the CSV file will be created.</param>
        public static void ExportRevisions(string inputDocPath, string outputCsvPath)
        {
            // Load the document (lifecycle rule: use provided load method)
            Document doc = new Document(inputDocPath);

            // Ensure the document contains revisions before proceeding
            if (!doc.HasRevisions)
            {
                Console.WriteLine("The document does not contain any revisions.");
                return;
            }

            // Open a StreamWriter for the CSV output (will create/overwrite the file)
            using (StreamWriter writer = new StreamWriter(outputCsvPath, false))
            {
                // Write CSV header
                writer.WriteLine("Author,DateTime,RevisionType,RevisionText");

                // Iterate over each revision in the document
                foreach (Revision rev in doc.Revisions)
                {
                    // Gather required fields
                    string author = rev.Author ?? string.Empty;
                    string dateTime = rev.DateTime.ToString("o"); // ISO 8601 format
                    string revType = rev.RevisionType.ToString();

                    // Get the text of the node that the revision is attached to.
                    // For safety, check for null (some revision types may not have a parent node).
                    string revText = rev.ParentNode != null ? rev.ParentNode.GetText().Replace("\r", " ").Replace("\n", " ") : string.Empty;

                    // Escape commas and quotes for CSV compliance
                    author = EscapeCsvField(author);
                    dateTime = EscapeCsvField(dateTime);
                    revType = EscapeCsvField(revType);
                    revText = EscapeCsvField(revText);

                    // Write the CSV line
                    writer.WriteLine($"{author},{dateTime},{revType},{revText}");
                }
            }

            // Demonstrate a save call – not required for export but kept to illustrate lifecycle usage.
            doc.Save(Path.ChangeExtension(outputCsvPath, ".docx"));
        }

        /// <summary>
        /// Escapes a CSV field by surrounding it with quotes if it contains a comma, quote or newline.
        /// Any existing quotes are doubled per CSV specification.
        /// </summary>
        private static string EscapeCsvField(string field)
        {
            if (field.Contains("\""))
                field = field.Replace("\"", "\"\"");

            if (field.Contains(",") || field.Contains("\r") || field.Contains("\n") || field.Contains("\""))
                field = $"\"{field}\"";

            return field;
        }
    }

    /// <summary>
    /// Console entry point that forwards arguments to <see cref="RevisionExporter.ExportRevisions"/>.
    /// </summary>
    public static class Program
    {
        /// <summary>
        /// Expected arguments: <c>inputDocPath outputCsvPath</c>
        /// </summary>
        public static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: RevisionExportUtility <input-doc-path> <output-csv-path>");
                return;
            }

            string inputDocPath = args[0];
            string outputCsvPath = args[1];

            try
            {
                RevisionExporter.ExportRevisions(inputDocPath, outputCsvPath);
                Console.WriteLine($"Revisions exported to '{outputCsvPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error exporting revisions: {ex.Message}");
            }
        }
    }
}
