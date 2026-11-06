using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder for inserting paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert sample paragraphs.
        builder.Writeln("First paragraph. This is a sample text.");
        builder.Writeln("Second paragraph. Another line of text.");
        builder.Writeln("Third paragraph. Yet another line.");

        // Save the document to a memory stream as plain text.
        using (MemoryStream txtStream = new MemoryStream())
        {
            TxtSaveOptions saveOptions = new TxtSaveOptions();
            doc.Save(txtStream, saveOptions);

            // Convert the stream to a string using the encoding defined in the save options (default UTF-8).
            string txtContent = Encoding.UTF8.GetString(txtStream.ToArray());

            // Split the text into lines and prefix each line with its line number.
            string[] lines = txtContent.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            StringBuilder numberedContent = new StringBuilder();
            for (int i = 0; i < lines.Length; i++)
            {
                // Skip adding a line number for an empty line at the end that may result from splitting.
                if (i == lines.Length - 1 && string.IsNullOrEmpty(lines[i]))
                    continue;

                numberedContent.AppendLine($"{i + 1}: {lines[i]}");
            }

            // Write the numbered text to the output file.
            string outputPath = "OutputWithLineNumbers.txt";
            File.WriteAllText(outputPath, numberedContent.ToString(), Encoding.UTF8);
        }
    }
}
