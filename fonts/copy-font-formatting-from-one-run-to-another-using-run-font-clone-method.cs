using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add first run with specific font formatting.
        builder.Font.Name = "Arial";
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.Writeln("First run text");

        // Retrieve the first run.
        Run run1 = (Run)doc.GetChild(NodeType.Run, 0, true);

        // Add second run with default formatting.
        builder.Writeln("Second run text");

        // Retrieve the second run.
        Run run2 = (Run)doc.GetChild(NodeType.Run, 1, true);

        // Copy font formatting from run1 to run2 by assigning individual properties.
        run2.Font.Name = run1.Font.Name;
        run2.Font.Size = run1.Font.Size;
        run2.Font.Bold = run1.Font.Bold;

        // Validate that the font properties were copied.
        bool isCopySuccessful = run2.Font.Name == run1.Font.Name &&
                                run2.Font.Size == run1.Font.Size &&
                                run2.Font.Bold == run1.Font.Bold;

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Ensure the output file exists.
        if (File.Exists(outputPath) && isCopySuccessful)
        {
            Console.WriteLine("Font formatting copied successfully and file saved.");
        }
        else
        {
            Console.WriteLine("Failed to copy font formatting or save the file.");
        }
    }
}
