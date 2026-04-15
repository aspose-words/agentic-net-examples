using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Prepare folders for input and output documents.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample DOCX files with simple equations.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            builder.Writeln($"Sample document {i}");
            // Insert a simple fraction equation using the deterministic EQ-field bootstrap workflow.
            InsertEquation(builder, @"\f(1,2)");
            builder.Writeln(); // Add a blank line after the equation.

            string inputPath = Path.Combine(inputDir, $"Doc{i}.docx");
            sampleDoc.Save(inputPath, SaveFormat.Docx);
        }

        // Process each document: standardize OfficeMath justification.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(filePath);

            // Retrieve all OfficeMath nodes in the document.
            NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
            foreach (Node node in mathNodes)
            {
                OfficeMath officeMath = (OfficeMath)node;
                // Apply only to top‑level equations (OMathPara).
                if (officeMath.MathObjectType == MathObjectType.OMathPara)
                {
                    // Ensure the equation is in display mode before setting justification.
                    officeMath.DisplayType = OfficeMathDisplayType.Display;
                    officeMath.Justification = OfficeMathJustification.Center;
                }
            }

            // Save the modified document to the output folder.
            string fileName = Path.GetFileName(filePath);
            string outputPath = Path.Combine(outputDir, fileName);
            doc.Save(outputPath, SaveFormat.Docx);

            // Validate that the file was saved.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to save output file: {outputPath}");
        }

        // Optional: indicate completion.
        Console.WriteLine("Processing completed. Output files are located in:");
        Console.WriteLine(outputDir);
    }

    // Helper method that inserts an EQ field, converts it to a real OfficeMath node, and returns it.
    private static OfficeMath InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Write the equation arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);
        // Return to the field start position.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start and remove the field.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();
        }

        return officeMath;
    }
}
