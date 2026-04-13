using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several paragraphs with placeholder text.
        for (int i = 0; i < 6; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Loop through all paragraphs and set font color based on the paragraph index.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        for (int i = 0; i < paragraphs.Count; i++)
        {
            Paragraph para = (Paragraph)paragraphs[i];

            // Determine the color: even index -> Red, odd index -> Blue.
            Aspose.Drawing.Color aspColor = (i % 2 == 0) ? Aspose.Drawing.Color.Red : Aspose.Drawing.Color.Blue;

            // Convert Aspose.Drawing.Color to System.Drawing.Color as required by the Font.Color property.
            System.Drawing.Color sysColor = System.Drawing.Color.FromArgb(aspColor.ToArgb());

            // Apply the color to all runs within the paragraph.
            foreach (Run run in para.Runs)
            {
                run.Font.Color = sysColor;

                // Validation: ensure the color was set correctly.
                if (run.Font.Color.ToArgb() != sysColor.ToArgb())
                {
                    throw new InvalidOperationException("Font color assignment validation failed.");
                }
            }
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DynamicFontColor.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not created.", outputPath);
        }

        // Optional: inform the user that the process completed.
        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
