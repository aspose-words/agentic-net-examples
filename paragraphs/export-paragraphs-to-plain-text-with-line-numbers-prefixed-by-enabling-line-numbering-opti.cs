using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several paragraphs.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");
        builder.Writeln("Third paragraph.");

        // Enable line numbering for the section.
        // The numbers will be used later when we export the text.
        var pageSetup = doc.FirstSection.PageSetup;
        pageSetup.LineStartingNumber = 1;               // Start numbering at 1.
        pageSetup.LineNumberCountBy = 1;                // Increment by 1.
        pageSetup.LineNumberRestartMode = LineNumberRestartMode.Continuous;
        pageSetup.LineNumberDistanceFromText = 0;       // Default distance.

        // Save the document to a temporary plain‑text file using TxtSaveOptions.
        // This demonstrates using SaveOptions while saving.
        TxtSaveOptions txtOptions = new TxtSaveOptions();
        string tempTxtPath = "temp.txt";
        doc.Save(tempTxtPath, txtOptions);

        // Export paragraphs to a new text file with line numbers prefixed.
        string finalTxtPath = "ExportedWithLineNumbers.txt";
        using (StreamWriter writer = new StreamWriter(finalTxtPath))
        {
            // Retrieve all paragraph nodes.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            int lineNumber = 1;

            foreach (Paragraph para in paragraphs)
            {
                // Get the paragraph text without the trailing paragraph break.
                string paraText = para.GetText().TrimEnd('\r', '\n');
                // Write the line number and paragraph text.
                writer.WriteLine($"{lineNumber}: {paraText}");
                lineNumber++;
            }
        }

        // Output the location of the generated file.
        Console.WriteLine($"Paragraphs exported with line numbers to: {Path.GetFullPath(finalTxtPath)}");
    }
}
