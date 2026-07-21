using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a paragraph that contains two runs with different formatting.
        builder.Font.Bold = true;               // First run – bold.
        builder.Write("Hello ");
        builder.Font.Bold = false;              // Second run – regular.
        builder.Write("World");
        builder.Writeln("!");                   // End the paragraph.

        // The paragraph we just created.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Replace the word "World" with "Aspose" while keeping each run's formatting.
        foreach (Run run in paragraph.Runs)
        {
            if (run.Text.Contains("World"))
            {
                run.Text = run.Text.Replace("World", "Aspose");
            }
        }

        // Save the result to a file in the current directory.
        doc.Save("Output.docx");
    }
}
