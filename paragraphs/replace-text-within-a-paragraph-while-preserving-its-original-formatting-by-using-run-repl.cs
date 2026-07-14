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
        // First run: normal text.
        builder.Font.Bold = false;
        builder.Write("Hello ");

        // Second run: bold text that we will replace.
        builder.Font.Bold = true;
        builder.Write("World!");

        // Replace the word "World" with "Aspose" while preserving the original formatting of each run.
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            foreach (Run run in paragraph.Runs)
            {
                if (run.Text.Contains("World"))
                {
                    // Only the text inside this run is changed; its formatting (bold) remains intact.
                    run.Text = run.Text.Replace("World", "Aspose");
                }
            }
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
