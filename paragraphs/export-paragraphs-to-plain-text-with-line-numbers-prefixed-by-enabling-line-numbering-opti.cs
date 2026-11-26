using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a DocumentBuilder for editing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph with more text.");
        builder.Writeln("Third paragraph.");

        // Save the document to a memory stream as plain text (no line numbers option available).
        using (MemoryStream txtStream = new MemoryStream())
        {
            TxtSaveOptions txtOptions = new TxtSaveOptions();
            doc.Save(txtStream, txtOptions);

            // Convert the stream to a string.
            txtStream.Position = 0;
            string txtContent;
            using (StreamReader reader = new StreamReader(txtStream))
            {
                txtContent = reader.ReadToEnd();
            }

            // Prefix each line with its line number.
            string[] lines = txtContent.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            for (int i = 0; i < lines.Length; i++)
            {
                lines[i] = $"{i + 1}: {lines[i]}";
            }

            // Write the numbered lines to the output file.
            string outputPath = "output.txt";
            File.WriteAllLines(outputPath, lines);

            Console.WriteLine($"Document saved to {outputPath}");
        }
    }
}
