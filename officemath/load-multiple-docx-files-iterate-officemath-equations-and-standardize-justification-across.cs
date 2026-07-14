using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathJustificationBatch
{
    public static void Main()
    {
        // Prepare input and output folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample DOCX files with OfficeMath equations.
        CreateSampleDocument(Path.Combine(inputDir, "Sample1.docx"));
        CreateSampleDocument(Path.Combine(inputDir, "Sample2.docx"));

        // Process each document: standardize justification of top‑level equations.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(filePath);

            // Enumerate all OfficeMath nodes.
            var officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
            foreach (OfficeMath om in officeMathNodes.OfType<OfficeMath>())
            {
                // Target only top‑level equations (MathObjectType.OMathPara).
                if (om.MathObjectType == MathObjectType.OMathPara)
                {
                    // Display type must be set before justification.
                    om.DisplayType = OfficeMathDisplayType.Display;
                    om.Justification = OfficeMathJustification.Center;
                }
            }

            // Save the modified document.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);

            // Validate that the file was saved.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to save the processed document: {outputPath}");
        }
    }

    // Creates a simple document containing a few OfficeMath equations using the EQ‑field bootstrap workflow.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Sample document with equations:");

        // Insert a fraction equation: 1/2
        InsertEquation(builder, @"\f(1,2)");

        // Insert an integral equation.
        InsertEquation(builder, @"\i");

        // Save the sample document.
        doc.Save(filePath);

        // Validate that the file was created.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create the sample document: {filePath}");
    }

    // Inserts an EQ field, converts it to a real OfficeMath node, and removes the original field.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return to the field start's parent node.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath before the field start and remove the field.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();
        }

        // Add a paragraph break after the equation for readability.
        builder.Writeln();
    }
}
