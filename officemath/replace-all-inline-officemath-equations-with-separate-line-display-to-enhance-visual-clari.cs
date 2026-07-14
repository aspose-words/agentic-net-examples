using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class ReplaceInlineOfficeMath
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a paragraph that contains several inline equations using EQ fields.
        builder.Writeln("Paragraph with inline equations:");
        builder.Writeln();

        // First inline equation.
        builder.Write("The fraction ");
        InsertFieldEQ(builder, @"\f(1,2)"); // 1/2
        builder.Write(" appears here. ");

        // Second inline equation.
        builder.Write("A radical: ");
        InsertFieldEQ(builder, @"\r(3,x)"); // cube root of x
        builder.Write(" is shown. ");

        // Third inline equation.
        builder.Write("Summation: ");
        InsertFieldEQ(builder, @"\i \su(n=1,5,n)"); // sum from n=1 to 5 of n
        builder.Writeln(".");

        // Convert all EQ fields to real OfficeMath objects.
        foreach (FieldEQ field in doc.Range.Fields.OfType<FieldEQ>().ToList())
        {
            OfficeMath officeMath = field.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                field.Start.ParentNode.InsertBefore(officeMath, field.Start);
                // Remove the original field.
                field.Remove();
            }
        }

        // Replace all inline top‑level OfficeMath equations with display equations.
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true).OfType<OfficeMath>())
        {
            if (om.MathObjectType == MathObjectType.OMathPara &&
                om.DisplayType == OfficeMathDisplayType.Inline)
            {
                // Change to display mode and left justification.
                om.DisplayType = OfficeMathDisplayType.Display;
                om.Justification = OfficeMathJustification.Left;
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved.");

        // End of example.
    }

    // Helper method that inserts an EQ field with the specified arguments.
    private static void InsertFieldEQ(DocumentBuilder builder, string args)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Move to the field separator and write the arguments.
        builder.MoveTo(field.Separator);
        builder.Write(args);
        // Return the builder to the field's start parent node for further writing.
        builder.MoveTo(field.Start.ParentNode);
        // Insert a space after the equation for readability.
        builder.Write(" ");
    }
}
