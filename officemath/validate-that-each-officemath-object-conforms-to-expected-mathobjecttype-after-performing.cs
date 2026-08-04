using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Fields;

public class OfficeMathBulkValidation
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "OfficeMathBulkUpdate.docx");

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Simple EQ field arguments that reliably convert to OfficeMath.
        List<string> equations = new List<string>
        {
            @"\f(1,2)",   // fraction 1/2
            @"\r(3,x)",   // cube root of x
            @"\i"         // integral symbol
        };

        // Insert each equation, convert the field to a real OfficeMath node, and remove the field.
        foreach (string eq in equations)
        {
            // Insert an empty EQ field.
            FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

            // Write the equation arguments into the field separator.
            builder.MoveTo(field.Separator);
            builder.Write(eq);

            // Return the builder to the field start's parent (the paragraph).
            builder.MoveTo(field.Start.ParentNode);

            // Convert the field to OfficeMath.
            OfficeMath officeMath = field.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                field.Start.ParentNode.InsertBefore(officeMath, field.Start);
                // Remove the original field.
                field.Remove();
            }

            // Start a new paragraph for the next equation.
            builder.Writeln();
        }

        // Bulk update: set display type and justification for all top‑level equations.
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                om.DisplayType = OfficeMathDisplayType.Display;
                om.Justification = OfficeMathJustification.Left;
            }
        }

        // Validation: ensure every OfficeMath node is of the expected type (OMathPara for top‑level).
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            if (om.MathObjectType != MathObjectType.OMathPara)
            {
                throw new Exception($"Unexpected MathObjectType detected: {om.MathObjectType}");
            }
        }

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output document was not saved correctly.");
        }

        // Indicate successful completion.
        Console.WriteLine("OfficeMath bulk update and validation completed successfully.");
    }
}
