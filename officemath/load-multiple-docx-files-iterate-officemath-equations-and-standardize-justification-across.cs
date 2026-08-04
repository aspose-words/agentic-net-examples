using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files with simple equations.
        for (int i = 1; i <= 2; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample document {i}");

            // Insert a safe EQ field that will be converted to a real OfficeMath node.
            InsertEquation(builder, @"\f(1,2)"); // Simple fraction equation.

            string inputPath = Path.Combine(inputFolder, $"Doc{i}.docx");
            sampleDoc.Save(inputPath, SaveFormat.Docx);
        }

        // Process each DOCX: standardize justification of top‑level OfficeMath nodes.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(filePath);

            var topLevelMath = doc.GetChildNodes(NodeType.OfficeMath, true)
                                  .OfType<OfficeMath>()
                                  .Where(om => om.MathObjectType == MathObjectType.OMathPara);

            foreach (OfficeMath om in topLevelMath)
            {
                // Ensure the equation is displayed on its own line before setting justification.
                om.DisplayType = OfficeMathDisplayType.Display;
                om.Justification = OfficeMathJustification.Center;
            }

            // Save the modified document.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath, SaveFormat.Docx);

            // Validation: reload and verify justification.
            Document verifyDoc = new Document(outputPath);
            OfficeMath firstMath = verifyDoc.GetChildNodes(NodeType.OfficeMath, true)
                                           .OfType<OfficeMath>()
                                           .FirstOrDefault(om => om.MathObjectType == MathObjectType.OMathPara);

            if (firstMath == null)
                throw new InvalidOperationException($"No top‑level OfficeMath found in '{outputPath}'.");

            if (firstMath.Justification != OfficeMathJustification.Center)
                throw new InvalidOperationException($"Justification was not set correctly in '{outputPath}'.");
        }

        // Ensure output files exist.
        if (!Directory.GetFiles(outputFolder, "*.docx").Any())
            throw new InvalidOperationException("No output documents were created.");
    }

    // Inserts an EQ field, converts it to a real OfficeMath node, and removes the field.
    private static void InsertEquation(DocumentBuilder builder, string eqArgs)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the equation arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArgs);

        // Update the field so that Aspose.Words parses the EQ code.
        field.Update();

        // Convert the field to OfficeMath.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath before the field and remove the field.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        field.Remove();
    }
}
