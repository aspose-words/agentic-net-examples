using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with several inline equations.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few paragraphs, each containing an inline equation.
        for (int i = 1; i <= 5; i++)
        {
            // Start a new paragraph with some leading text.
            builder.Writeln($"Paragraph {i} with an inline equation:");

            // Insert an EQ field that will be converted to a real OfficeMath object.
            FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
            // Move to the field separator and write a simple fraction equation.
            builder.MoveTo(eqField.Separator);
            builder.Write(@"\f(1,2)");
            // Return the builder to the paragraph that contains the field.
            builder.MoveTo(eqField.Start.ParentNode);

            // Convert the EQ field to OfficeMath.
            OfficeMath officeMath = eqField.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start and remove the field.
                eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
                eqField.Remove();

                // Ensure the equation is displayed inline initially.
                officeMath.DisplayType = OfficeMathDisplayType.Inline;
            }

            // Add a blank line after each equation for readability.
            builder.Writeln();
        }

        // Save the intermediate document (optional, demonstrates the bootstrap workflow).
        string intermediatePath = Path.Combine(outputDir, "Report_With_Inline_Equations.docx");
        doc.Save(intermediatePath, SaveFormat.Docx);

        // Now change all top‑level OfficeMath equations to display on a separate line.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath math in mathNodes)
        {
            // Only modify top‑level equations (MathObjectType == OMathPara).
            if (math.MathObjectType == MathObjectType.OMathPara)
            {
                math.DisplayType = OfficeMathDisplayType.Display;
                math.Justification = OfficeMathJustification.Left;
            }
        }

        // Save the final document.
        string finalPath = Path.Combine(outputDir, "Report_With_Display_Equations.docx");
        doc.Save(finalPath, SaveFormat.Docx);

        // Simple validation that the output file was created.
        if (!File.Exists(finalPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
