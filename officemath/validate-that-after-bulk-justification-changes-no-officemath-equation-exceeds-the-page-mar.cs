using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class OfficeMathJustificationValidator
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three sample equations using the deterministic EQ-field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");          // Simple fraction 1/2
        InsertEquation(builder, @"\r(3,x)");          // Cube root of x
        InsertEquation(builder, @"\i \su(n=1,5,n)"); // Integral with summation

        // Convert each inserted EQ field to a real OfficeMath node.
        foreach (FieldEQ field in doc.Range.Fields.OfType<FieldEQ>())
        {
            OfficeMath officeMath = field.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start and remove the field.
                field.Start.ParentNode.InsertBefore(officeMath, field.Start);
                field.Remove();
            }
        }

        // Bulk change: set justification of all top‑level OfficeMath paragraphs to CenterGroup.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath math in mathNodes.Cast<OfficeMath>()
                                            .Where(m => m.MathObjectType == MathObjectType.OMathPara))
        {
            // Display type must be set before justification when changing to Inline/Display.
            math.DisplayType = OfficeMathDisplayType.Display;
            math.Justification = OfficeMathJustification.CenterGroup;
        }

        // Validate that no equation exceeds the page margins.
        double pageWidth = doc.FirstSection.PageSetup.PageWidth;
        double leftMargin = doc.FirstSection.PageSetup.LeftMargin;
        double rightMargin = doc.FirstSection.PageSetup.RightMargin;
        double usableWidth = pageWidth - leftMargin - rightMargin; // points

        foreach (OfficeMath math in mathNodes.Cast<OfficeMath>()
                                            .Where(m => m.MathObjectType == MathObjectType.OMathPara))
        {
            // Get the rendered width of the equation.
            double equationWidth = math.GetMathRenderer().SizeInPoints.Width;

            if (equationWidth > usableWidth)
                throw new InvalidOperationException(
                    $"Equation exceeds page margins: width {equationWidth} pts, usable {usableWidth} pts.");
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OfficeMathJustificationValidated.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        Console.WriteLine("Validation completed successfully. Document saved to:");
        Console.WriteLine(outputPath);
    }

    // Helper that inserts an EQ field with the given argument string,
    // then moves the builder back to the paragraph after the field.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Write the arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);
        // Return the builder to the paragraph after the field.
        builder.MoveTo(field.Start.ParentNode);
        builder.InsertParagraph();
    }
}
