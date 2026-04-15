using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathBatchInserter
{
    // Predefined EQ field argument that creates a simple fraction 1/2.
    private const string EquationArgument = @"\f(1,2)";

    public static void Main()
    {
        // Prepare input and output folders.
        string baseDir = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory);
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a sample source document if none exist.
        string samplePath = Path.Combine(inputDir, "Sample.docx");
        if (!File.Exists(samplePath))
        {
            CreateSampleDocument(samplePath);
        }

        // Process each .docx file in the input folder.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(filePath);
            InsertEquationIntoEveryParagraph(doc);
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);
            // Simple validation that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to save output file: {outputPath}");
        }
    }

    // Creates a simple document with a few paragraphs of placeholder text.
    private static void CreateSampleDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Paragraph {i}: This is sample text.");
        }

        doc.Save(path);
    }

    // Inserts the predefined OfficeMath equation into each paragraph of the given document.
    private static void InsertEquationIntoEveryParagraph(Document doc)
    {
        // Collect all paragraphs first to avoid modifying the collection while iterating.
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                            .Cast<Paragraph>()
                            .ToList();

        foreach (Paragraph paragraph in paragraphs)
        {
            // Use a new builder for each insertion to avoid cursor conflicts.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(paragraph.LastChild ?? paragraph);

            // Insert an EQ field.
            FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
            // Write the equation argument into the field separator.
            builder.MoveTo(eqField.Separator);
            builder.Write(EquationArgument);
            // Return the builder to the paragraph context.
            builder.MoveTo(eqField.Start.ParentNode);

            // Convert the EQ field to a real OfficeMath object.
            OfficeMath officeMath = eqField.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start node.
                eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
                // Remove the original field so only the OfficeMath remains.
                eqField.Remove();

                // Optional: set display type to display on its own line and left‑justified.
                officeMath.DisplayType = OfficeMathDisplayType.Display;
                officeMath.Justification = OfficeMathJustification.Left;
            }
        }
    }
}
