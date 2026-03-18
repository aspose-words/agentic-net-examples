using System;
using System.IO;
using Aspose.Words;

class ExportCommentsToCsv
{
    static void Main()
    {
        // Load the DOCX file using the Aspose.Words Document constructor (load rule)
        Document doc = new Document("input.docx");

        // Create a CSV file and write the header row
        using (StreamWriter writer = new StreamWriter("comments.csv"))
        {
            writer.WriteLine("Author,Date,Text");

            // Retrieve all comment nodes from the document (including replies)
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            foreach (Comment comment in comments)
            {
                // Extract required fields
                string author = comment.Author;
                string date   = comment.DateTime.ToString("o"); // ISO 8601 format
                string text   = comment.GetText().Trim();

                // Escape each field for CSV compliance
                string authorEsc = EscapeCsv(author);
                string dateEsc   = EscapeCsv(date);
                string textEsc   = EscapeCsv(text);

                // Write a CSV line
                writer.WriteLine($"{authorEsc},{dateEsc},{textEsc}");
            }
        }
    }

    // Escapes a string for CSV: encloses in quotes if needed and doubles internal quotes
    private static string EscapeCsv(string value)
    {
        if (value == null) return string.Empty;

        bool mustQuote = value.Contains(",") || value.Contains("\"") || value.Contains("\r") || value.Contains("\n");
        if (mustQuote)
        {
            string escaped = value.Replace("\"", "\"\"");
            return $"\"{escaped}\"";
        }
        return value;
    }
}
