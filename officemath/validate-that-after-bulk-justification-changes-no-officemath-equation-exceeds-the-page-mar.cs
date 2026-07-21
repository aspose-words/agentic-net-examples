using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Rendering;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several simple equations using the deterministic EQ‑field bootstrap workflow.
        string[] eqArguments = new[]
        {
            @"\f(1,2)",          // fraction 1/2
            @"\r(3,x)",          // cube root of x
            @"\i \su(n=1,5,n)", // integral with summation
            @"\a \co2 \vs1 \hs1(1,2,3,4)", // simple array
            @"\f(3,4)"           // fraction 3/4
        };

        foreach (string args in eqArguments)
        {
            // Insert an EQ field.
            FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
            // Write the arguments into the field separator.
            builder.MoveTo(field.Separator);
            builder.Write(args);
            // Return the builder to the field start position.
            builder.MoveTo(field.Start.ParentNode);

            // Convert the field to a real OfficeMath object.
            OfficeMath officeMath = field.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                field.Start.ParentNode.InsertBefore(officeMath, field.Start);
                // Remove the original field.
                field.Remove();
            }

            // Add a paragraph break after each equation.
            builder.Writeln();
        }

        // Ensure the document has at least one section (it does by default) and set explicit margins.
        Section section = doc.FirstSection;
        section.PageSetup.LeftMargin = ConvertUtil.InchToPoint(1.0);
        section.PageSetup.RightMargin = ConvertUtil.InchToPoint(1.0);
        section.PageSetup.TopMargin = ConvertUtil.InchToPoint(1.0);
        section.PageSetup.BottomMargin = ConvertUtil.InchToPoint(1.0);

        // Bulk justification change: set each top‑level OfficeMath node to Display mode and CenterGroup justification.
        var officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true)
                                 .Cast<OfficeMath>()
                                 .Where(om => om.MathObjectType == MathObjectType.OMathPara);

        foreach (OfficeMath om in officeMathNodes)
        {
            // Display type must be set before justification.
            om.DisplayType = OfficeMathDisplayType.Display;
            om.Justification = OfficeMathJustification.CenterGroup;
        }

        // Validation: ensure no equation exceeds the printable width (page width minus margins).
        double pageWidth = section.PageSetup.PageWidth;
        double leftMargin = section.PageSetup.LeftMargin;
        double rightMargin = section.PageSetup.RightMargin;
        double maxContentWidth = pageWidth - leftMargin - rightMargin;
        const double tolerance = 0.5; // points tolerance for rounding differences.

        foreach (OfficeMath om in officeMathNodes)
        {
            OfficeMathRenderer renderer = om.GetMathRenderer();
            double equationWidth = renderer.SizeInPoints.Width;

            if (equationWidth > maxContentWidth + tolerance)
            {
                throw new InvalidOperationException(
                    $"Equation exceeds page margins. Width: {equationWidth} pts, Max allowed: {maxContentWidth} pts.");
            }
        }

        // Save the resulting document.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "JustifiedEquations.docx");
        doc.Save(outputPath);

        // Indicate successful completion.
        Console.WriteLine("Document saved successfully. All equations fit within page margins.");
    }

    // Helper class for unit conversions.
    private static class ConvertUtil
    {
        private const double PointsPerInch = 72.0;
        public static double InchToPoint(double inches) => inches * PointsPerInch;
    }
}
