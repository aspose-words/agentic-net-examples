using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class ReplaceInlineOfficeMath
{
    public static void Main()
    {
        // Define folders for input and output documents.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        string inputPath = Path.Combine(artifactsDir, "Input.docx");
        string outputPath = Path.Combine(artifactsDir, "Output.docx");

        // -----------------------------------------------------------------
        // Step 1: Create a sample document with several inline OfficeMath equations.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with some text.
        builder.Writeln("Below are several inline equations:");

        // Helper to insert an EQ field, convert it to OfficeMath, and keep it inline.
        void InsertInlineEquation(string eqArgs)
        {
            // Insert an EQ field.
            FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
            // Write the EQ arguments.
            builder.MoveTo(field.Separator);
            builder.Write(eqArgs);
            // Return to the paragraph.
            builder.MoveTo(field.Start.ParentNode);
            // Convert the field to a real OfficeMath object.
            OfficeMath officeMath = field.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath before the field start.
                field.Start.ParentNode.InsertBefore(officeMath, field.Start);
                // Ensure the equation is initially inline.
                officeMath.DisplayType = OfficeMathDisplayType.Inline;
                // Remove the original field.
                field.Remove();
            }
            // Add a space after the equation for readability.
            builder.Write(" ");
        }

        // Insert a few simple equations.
        InsertInlineEquation(@"\f(1,2)");          // Fraction 1/2
        InsertInlineEquation(@"\i");               // Integral symbol
        InsertInlineEquation(@"\r(3,x)");          // Cube root of x

        // Finish the paragraph.
        builder.Writeln();

        // Save the sample document.
        doc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Step 2: Load the document and replace inline equations with display equations.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // Find all top‑level OfficeMath nodes.
        NodeCollection officeMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in officeMathNodes)
        {
            // Target only top‑level equations (OMathPara) that are currently inline.
            if (om.MathObjectType == MathObjectType.OMathPara &&
                om.DisplayType == OfficeMathDisplayType.Inline)
            {
                // Change to display mode and left‑justify.
                om.DisplayType = OfficeMathDisplayType.Display;
                om.Justification = OfficeMathJustification.Left;
            }
        }

        // Save the modified document.
        loadedDoc.Save(outputPath, SaveFormat.Docx);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
