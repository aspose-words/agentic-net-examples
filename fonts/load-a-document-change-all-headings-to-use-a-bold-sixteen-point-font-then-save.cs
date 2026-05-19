using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary input and final output documents.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "Input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document that contains headings.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Add a Heading 1.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Sample Heading 1");

        // Add a Heading 2.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Sample Heading 2");

        // Add a normal paragraph (should remain unchanged).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Normal paragraph text.");

        // Save the sample document to disk.
        sampleDoc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the existing document.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Change all heading styles to use a bold 16‑point font.
        // -----------------------------------------------------------------
        // Iterate through all built‑in heading style identifiers (Heading1 … Heading9).
        for (int id = (int)StyleIdentifier.Heading1; id <= (int)StyleIdentifier.Heading9; id++)
        {
            StyleIdentifier styleId = (StyleIdentifier)id;
            // Retrieve the style; it may be null if the document does not contain it.
            Aspose.Words.Style headingStyle = doc.Styles[styleId];
            if (headingStyle != null)
            {
                // Apply the desired formatting to the style's font.
                headingStyle.Font.Size = 16;          // sixteen‑point size
                headingStyle.Font.Bold = true;        // bold
            }
        }

        // -----------------------------------------------------------------
        // 4. Save the modified document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);

        // -----------------------------------------------------------------
        // 5. Simple validation that the output file was created.
        // -----------------------------------------------------------------
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully to: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
