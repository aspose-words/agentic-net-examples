using System;
using Aspose.Words;

public class BuildToken
{
    // Simple token that can be set to request interruption.
    public bool CancelRequested { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a token that will be passed to the builder.
        BuildToken token = new BuildToken();

        // Simulate a condition that requests cancellation after a few iterations.
        // In a real scenario this could be set from another thread or based on user input.
        int cancelAfter = 5;

        // Build the document while periodically checking the token.
        BuildDocument(doc, token, cancelAfter);

        // Save the resulting document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }

    private static void BuildDocument(Document doc, BuildToken token, int cancelAfter)
    {
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < 100; i++)
        {
            // Check the token before each operation.
            if (token.CancelRequested)
            {
                Console.WriteLine("Document building was interrupted by token.");
                break;
            }

            // Write a paragraph.
            builder.Writeln($"Paragraph {i + 1}");

            // Simulate a condition that triggers cancellation.
            if (i + 1 == cancelAfter)
            {
                token.CancelRequested = true;
            }
        }
    }
}
