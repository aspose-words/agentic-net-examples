using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Fields;

public class OfficeMathBulkUpdateExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Simple EQ field arguments for safe equation creation.
        string[] equations = new string[]
        {
            @"\f(1,2)",   // Fraction 1/2
            @"\r(3,x)",   // Cube root of x
            @"\i",        // Integral symbol
            @"\s \up5(Sup)", // Superscript
            @"\s \do5(Sub)"  // Subscript
        };

        // Insert each equation as a top‑level OfficeMath node.
        foreach (string eq in equations)
        {
            // Insert an EQ field.
            FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
            // Write the equation arguments into the field separator.
            builder.MoveTo(field.Separator);
            builder.Write(eq);
            // Return to the paragraph that contains the field.
            builder.MoveTo(field.Start.ParentNode);

            // Convert the EQ field to a real OfficeMath object.
            OfficeMath officeMath = field.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath before the field start and remove the field.
                field.Start.ParentNode.InsertBefore(officeMath, field.Start);
                field.Remove();
            }

            // Start a new paragraph for the next equation.
            builder.Writeln();
        }

        // Retrieve all OfficeMath nodes in the document.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Bulk update: set display type and justification for top‑level equations.
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                om.DisplayType = OfficeMathDisplayType.Display;
                om.Justification = OfficeMathJustification.Left;
            }
        }

        // Validation: ensure every OfficeMath node is of type OMathPara.
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType != MathObjectType.OMathPara)
                throw new InvalidOperationException("Unexpected MathObjectType detected.");
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OfficeMathBulkUpdate.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);

        Console.WriteLine("Document saved and validation succeeded: " + outputPath);
    }
}
