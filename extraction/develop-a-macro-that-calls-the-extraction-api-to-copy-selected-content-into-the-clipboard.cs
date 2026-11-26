using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with three paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph - not selected.");
        builder.Writeln("Second paragraph - will be copied to clipboard.");
        builder.Writeln("Third paragraph - not selected.");

        // Save the sample document locally.
        const string sourcePath = "sample.docx";
        doc.Save(sourcePath);

        // Load the document from the file system.
        Document loaded = new Document(sourcePath);

        // Locate the second paragraph (index 1) that we want to copy.
        Paragraph paragraphToCopy = loaded.FirstSection?.Body?.Paragraphs[1];
        if (paragraphToCopy == null)
            throw new InvalidOperationException("The target paragraph was not found.");

        // Extract the text of the selected paragraph.
        string extractedText = paragraphToCopy.GetText();

        // In a typical desktop environment we would place the text on the system clipboard.
        // Since this example runs in a console without UI libraries, we simulate the clipboard
        // by writing the extracted text to a deterministic output file.
        const string clipboardPath = "clipboard.txt";
        File.WriteAllText(clipboardPath, extractedText);

        // Verify that the simulated clipboard file was created and contains the expected text.
        if (!File.Exists(clipboardPath))
            throw new InvalidOperationException("Expected clipboard output file was not created.");

        string clipboardText = File.ReadAllText(clipboardPath);
        if (clipboardText != extractedText)
            throw new InvalidOperationException("Simulated clipboard content does not match the extracted text.");

        // The example completes successfully if no exception is thrown.
        Console.WriteLine("Paragraph extracted and written to '" + clipboardPath + "'.");
    }
}
