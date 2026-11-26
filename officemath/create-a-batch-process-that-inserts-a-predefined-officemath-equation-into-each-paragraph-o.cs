using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class OfficeMathBatchInsert
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "OutputDocs");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a sample DOCX with several paragraphs.
        string samplePath = Path.Combine(inputFolder, "Sample.docx");
        CreateSampleDocument(samplePath);

        // Process each DOCX in the input folder.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(filePath);
            InsertEquationIntoEachParagraph(doc);
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath, SaveFormat.Docx);

            // Validate the result.
            ValidateDocument(outputPath);
        }
    }

    // Creates a simple document with three paragraphs.
    private static void CreateSampleDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Paragraph {i}");
        }
        doc.Save(path, SaveFormat.Docx);
    }

    // Inserts a simple OfficeMath equation into every paragraph of the document.
    private static void InsertEquationIntoEachParagraph(Document doc)
    {
        // Gather paragraphs first to avoid modifying the collection while iterating.
        List<Paragraph> paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                                        .Cast<Paragraph>()
                                        .ToList();

        foreach (Paragraph paragraph in paragraphs)
        {
            // Use a builder tied to the same document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(paragraph);

            // Insert an EQ field (the bootstrap for a real OfficeMath object).
            Field field = builder.InsertField(FieldType.FieldEquation, true);
            FieldEQ fieldEq = field as FieldEQ;
            if (fieldEq == null)
                continue; // Safety check.

            // Write a simple EQ argument at the field separator.
            builder.MoveTo(fieldEq.Separator);
            builder.Write(@"\f(1,2)");

            // Update the field so that Aspose can convert it.
            field.Update();

            // Convert the field to a real OfficeMath node.
            OfficeMath officeMath = fieldEq.AsOfficeMath();
            if (officeMath == null)
                continue; // Conversion failed; skip this paragraph.

            // Insert the OfficeMath node before the field start node.
            Node fieldStart = fieldEq.Start;
            CompositeNode parent = fieldStart.ParentNode as CompositeNode;
            if (parent != null)
                parent.InsertBefore(officeMath, fieldStart);

            // Remove the original field, leaving only the real OfficeMath node.
            fieldEq.Remove();
        }
    }

    // Validates that each paragraph now contains exactly one top‑level OfficeMath node.
    private static void ValidateDocument(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("Output file was not created.", path);

        Document doc = new Document(path);

        // Count top‑level OfficeMath nodes (MathObjectType.OMathPara).
        int topLevelCount = doc.GetChildNodes(NodeType.OfficeMath, true)
                               .Cast<OfficeMath>()
                               .Count(om => om.MathObjectType == MathObjectType.OMathPara);

        // Count paragraphs (each should contain one equation).
        int paragraphCount = doc.GetChildNodes(NodeType.Paragraph, true).Count;

        if (topLevelCount != paragraphCount)
            throw new InvalidOperationException(
                $"Validation failed: expected {paragraphCount} top‑level equations, but found {topLevelCount}.");
    }
}
