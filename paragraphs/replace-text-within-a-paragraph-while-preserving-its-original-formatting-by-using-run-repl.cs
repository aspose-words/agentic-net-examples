using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a paragraph that contains two runs with different formatting.
        builder.Font.Name = "Arial";
        builder.Font.Size = 12;
        builder.Font.Bold = true;          // First run: bold.
        builder.Write("Hello ");
        builder.Font.Bold = false;
        builder.Font.Italic = true;        // Second run: italic.
        builder.Write("World");
        builder.Writeln();                 // End the paragraph.

        // Replace the text "Hello" with "Hi" while preserving the original formatting.
        // The Range.Replace method updates the text inside the existing run(s) without altering their formatting.
        doc.Range.Replace("Hello", "Hi");

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
