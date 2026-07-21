using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Path for the output document.
        const string outputPath = "ReportWithDisplayChanged.docx";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a title.
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.Writeln("Large Report with Inline Equations");
        // Reset font formatting to defaults.
        builder.Font.ClearFormatting();
        builder.Writeln();

        // Insert several inline equations using the deterministic EQ‑field bootstrap workflow.
        const int equationCount = 10; // Simulate a large report.
        for (int i = 0; i < equationCount; i++)
        {
            builder.Write($"Equation {i + 1} (inline): ");

            // Insert an EQ field.
            FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

            // Move to the field separator and write a simple fraction equation.
            builder.MoveTo(field.Separator);
            builder.Write(@"\f(1,2)");

            // Return the builder to the paragraph containing the field.
            builder.MoveTo(field.Start.ParentNode);

            // Convert the field to a real OfficeMath object.
            OfficeMath officeMath = field.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                field.Start.ParentNode.InsertBefore(officeMath, field.Start);

                // Ensure the equation is initially inline.
                officeMath.DisplayType = OfficeMathDisplayType.Inline;

                // Remove the original EQ field.
                field.Remove();
            }

            builder.Writeln(); // Separate equations with a line break.
        }

        // Change all top‑level OfficeMath equations to display (separate line) format.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (Node node in mathNodes)
        {
            OfficeMath om = (OfficeMath)node;
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                om.DisplayType = OfficeMathDisplayType.Display;
                om.Justification = OfficeMathJustification.Left;
            }
        }

        // Save the modified document.
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
