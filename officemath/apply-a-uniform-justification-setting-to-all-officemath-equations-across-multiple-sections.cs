using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class ApplyOfficeMathJustification
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a sample document with several sections, each containing a few equations.
        for (int sec = 1; sec <= 3; sec++)
        {
            // Add a heading for the section.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Section {sec}");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

            // Insert three simple equations in the current section.
            InsertEquation(builder, @"\f(1,2)");   // fraction 1/2
            InsertEquation(builder, @"\r(3,x)");   // cubic root of x
            InsertEquation(builder, @"\i \su(n=1,5,n)"); // integral with summation

            // Insert a section break after the last section.
            if (sec < 3)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Apply a uniform justification to all top‑level OfficeMath equations.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                // Display type must be set before justification.
                om.DisplayType = OfficeMathDisplayType.Display;
                om.Justification = OfficeMathJustification.Center;
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Output.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Simple validation – ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }

    // Helper that inserts an EQ field, converts it to a real OfficeMath node, and removes the field.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return the builder to the paragraph containing the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start and remove the field.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();
        }

        // Add a line break after the equation for readability.
        builder.Writeln();
    }
}
