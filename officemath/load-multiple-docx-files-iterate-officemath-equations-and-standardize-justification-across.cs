using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create sample input documents with OfficeMath equations.
        string inputDir = "InputDocs";
        Directory.CreateDirectory(inputDir);
        CreateSampleDocument(Path.Combine(inputDir, "Sample1.docx"));
        CreateSampleDocument(Path.Combine(inputDir, "Sample2.docx"));

        // Process each document: standardize justification of top‑level equations.
        string outputDir = "OutputDocs";
        Directory.CreateDirectory(outputDir);
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(filePath);

            // Find all OfficeMath nodes.
            NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
            foreach (Node node in mathNodes)
            {
                OfficeMath officeMath = (OfficeMath)node;
                // Apply changes only to top‑level equations (OMathPara).
                if (officeMath.MathObjectType == MathObjectType.OMathPara)
                {
                    // Display type must be set before justification.
                    officeMath.DisplayType = OfficeMathDisplayType.Display;
                    officeMath.Justification = OfficeMathJustification.Center;
                }
            }

            // Save the modified document.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);

            // Validate that the file was saved.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to save output file: {outputPath}");
        }

        // Indicate completion.
        Console.WriteLine("All documents processed and saved successfully.");
    }

    // Creates a simple document containing two equations using the deterministic EQ‑field bootstrap workflow.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Sample document with equations:");

        // First equation: a simple fraction.
        InsertEquation(builder, @"\f(1,2)");
        builder.Writeln("Text between equations.");

        // Second equation: a cubic root.
        InsertEquation(builder, @"\r(3,x)");

        doc.Save(filePath);
    }

    // Inserts an EQ field, converts it to a real OfficeMath node, and removes the original field.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return the builder to the paragraph containing the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field.
            field.Remove();
        }

        // Add a paragraph break after the equation for readability.
        builder.InsertParagraph();
    }
}
