using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample EQ field arguments for simple equations.
        string[] equations = new[]
        {
            @"\f(1,2)",               // Fraction 1/2
            @"\r(3,x)",               // Cube root of x
            @"\i \su(n=1,5,n)",       // Integral with summation
            @"\s \up8(Superscript)", // Superscript
            @"\s \do8(Subscript)"    // Subscript
        };

        // Insert each equation into the document using the deterministic EQ‑field bootstrap workflow.
        foreach (string eq in equations)
        {
            // Insert an EQ field.
            FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

            // Write the equation arguments into the field separator.
            builder.MoveTo(field.Separator);
            builder.Write(eq);

            // Return to the field start paragraph.
            builder.MoveTo(field.Start.ParentNode);

            // Convert the EQ field to a real OfficeMath object.
            OfficeMath officeMath = field.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                field.Start.ParentNode.InsertBefore(officeMath, field.Start);

                // Remove the original field.
                field.Remove();

                // Move the builder after the inserted OfficeMath and start a new paragraph for the next equation.
                builder.MoveTo(officeMath);
                builder.InsertParagraph();
            }
        }

        // Change all top‑level OfficeMath equations from inline to display mode.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath math in mathNodes)
        {
            if (math.MathObjectType == MathObjectType.OMathPara)
            {
                math.DisplayType = OfficeMathDisplayType.Display;
                math.Justification = OfficeMathJustification.Left;
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ReportWithDisplayMath.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
