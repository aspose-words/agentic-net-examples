using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathTypeDemo
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a fraction equation: \f(1,2)
        InsertEquation(builder, @"\f(1,2)");

        // Insert a radical equation: \r(3,x)
        InsertEquation(builder, @"\r(3,x)");

        // Save the document.
        string docPath = "OfficeMathDemo.docx";
        doc.Save(docPath);

        // Reload the document.
        Document loadedDoc = new Document(docPath);

        // Traverse OfficeMath nodes.
        NodeCollection officeMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        int index = 0;
        foreach (OfficeMath om in officeMathNodes)
        {
            string typeDescription;
            switch (om.MathObjectType)
            {
                case MathObjectType.Fraction:
                    typeDescription = "Fraction";
                    break;
                case MathObjectType.Radical:
                    typeDescription = "Radical";
                    break;
                default:
                    typeDescription = "Other (" + om.MathObjectType + ")";
                    break;
            }

            Console.WriteLine($"OfficeMath node {index}: {typeDescription}");
            index++;
        }

        // Validate presence of both fraction and radical.
        bool hasFraction = false;
        bool hasRadical = false;
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType == MathObjectType.Fraction) hasFraction = true;
            if (om.MathObjectType == MathObjectType.Radical) hasRadical = true;
        }

        if (!hasFraction || !hasRadical)
            throw new InvalidOperationException("Expected both fraction and radical OfficeMath nodes.");

        // Write a simple report file.
        string reportPath = "OfficeMathReport.txt";
        using (StreamWriter writer = new StreamWriter(reportPath))
        {
            writer.WriteLine("OfficeMath nodes analysis:");
            index = 0;
            foreach (OfficeMath om in officeMathNodes)
            {
                writer.WriteLine($"Node {index}: MathObjectType = {om.MathObjectType}");
                index++;
            }
        }

        // Verify report file creation.
        if (!File.Exists(reportPath))
            throw new FileNotFoundException("Report file was not created.", reportPath);
    }

    // Helper that inserts an EQ field, writes the equation argument, updates the field,
    // converts it to OfficeMath, inserts the OfficeMath node, and removes the original field.
    private static void InsertEquation(DocumentBuilder builder, string eqArgument)
    {
        // Insert an EQ field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator node and write the EQ argument.
        builder.MoveTo(field.Separator);
        builder.Write(eqArgument);

        // Update the field so that Aspose.Words can convert it to OfficeMath.
        field.Update();

        // Convert the field to OfficeMath.
        FieldEQ fieldEQ = field as FieldEQ;
        OfficeMath officeMath = fieldEQ?.AsOfficeMath();

        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            Node fieldStart = field.Start;
            fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);

            // Remove the original field.
            field.Remove();

            // Move the builder after the inserted OfficeMath and start a new line.
            builder.MoveTo(officeMath);
            builder.Writeln();
        }
        else
        {
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");
        }
    }
}
