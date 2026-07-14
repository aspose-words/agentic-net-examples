using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathBatchInsert
{
    // Predefined EQ field argument for a simple fraction 1/2.
    private const string EquationArgs = @"\f(1,2)";

    public static void Main()
    {
        // Prepare input and output folders.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Input");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample source documents.
        CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"));
        CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"));

        // Process each document in the input folder.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(filePath);

            // Insert the predefined equation into every paragraph.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paragraphs)
            {
                InsertEquationIntoParagraph(doc, para);
            }

            // Save the processed document to the output folder.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath);

            // Simple validation that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create output file: {outputPath}");
        }
    }

    // Creates a simple document with three paragraphs for demonstration.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Sample paragraph {i}.");
        }

        doc.Save(filePath);
    }

    // Inserts the predefined OfficeMath equation into the given paragraph.
    private static void InsertEquationIntoParagraph(Document doc, Paragraph para)
    {
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Position the builder at the end of the paragraph.
        builder.MoveTo(para);
        builder.Write(" "); // Optional space before the equation.

        // Insert an EQ field that will be converted to OfficeMath.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEQ = (FieldEQ)field;

        // Write the equation arguments into the field separator.
        builder.MoveTo(fieldEQ.Separator);
        builder.Write(EquationArgs);

        // Return the builder to the paragraph node.
        builder.MoveTo(fieldEQ.Start.ParentNode);

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = fieldEQ.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start and remove the field.
            fieldEQ.Start.ParentNode.InsertBefore(officeMath, fieldEQ.Start);
            fieldEQ.Remove();
        }
    }
}
