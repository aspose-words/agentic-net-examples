using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathTypeReporter
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Insert a fraction equation ----------
        // Insert an EQ field.
        Field fractionField = builder.InsertField(FieldType.FieldEquation, true);
        // Move to the field separator and write the EQ argument for a fraction.
        builder.MoveTo(fractionField.Separator);
        builder.Write(@"\f(1,2)"); // Represents the fraction 1/2.
        // Convert the field to a real OfficeMath node.
        OfficeMath fractionMath = ((FieldEQ)fractionField).AsOfficeMath();
        if (fractionMath != null)
        {
            // Insert the OfficeMath node before the field start and remove the field.
            fractionField.Start.ParentNode.InsertBefore(fractionMath, fractionField.Start);
            fractionField.Remove();
        }

        // Add a new paragraph for the next equation.
        builder.Writeln();

        // ---------- Insert a radical equation ----------
        Field radicalField = builder.InsertField(FieldType.FieldEquation, true);
        builder.MoveTo(radicalField.Separator);
        builder.Write(@"\r(3,x)"); // Represents the cubic root of x.
        OfficeMath radicalMath = ((FieldEQ)radicalField).AsOfficeMath();
        if (radicalMath != null)
        {
            radicalField.Start.ParentNode.InsertBefore(radicalMath, radicalField.Start);
            radicalField.Remove();
        }

        // Save the document to disk.
        const string outputPath = "OfficeMathTypes.docx";
        doc.Save(outputPath);
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output document.");

        // ---------- Enumerate OfficeMath nodes and report their types ----------
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        using (StreamWriter reportWriter = new StreamWriter("MathTypesReport.txt"))
        {
            for (int i = 0; i < mathNodes.Count; i++)
            {
                OfficeMath om = (OfficeMath)mathNodes[i];
                string typeDescription = GetMathObjectTypeDescription(om.MathObjectType);
                string line = $"OfficeMath node #{i + 1}: {om.MathObjectType} ({typeDescription})";
                Console.WriteLine(line);
                reportWriter.WriteLine(line);
            }
        }

        // Verify the report file was created.
        if (!File.Exists("MathTypesReport.txt"))
            throw new InvalidOperationException("Failed to create the report file.");
    }

    // Helper method to translate MathObjectType to a friendly description.
    private static string GetMathObjectTypeDescription(MathObjectType type)
    {
        return type switch
        {
            MathObjectType.Fraction => "Fraction",
            MathObjectType.Radical => "Radical",
            MathObjectType.OMathPara => "Paragraph (top‑level equation)",
            _ => "Other"
        };
    }
}
