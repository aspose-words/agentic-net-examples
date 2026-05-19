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

        // Insert a regular paragraph before the equations.
        builder.Writeln("Paragraph before equations.");

        // Insert three equations. The second one will be removed later.
        OfficeMath eq1 = InsertEquation(builder, @"\f(1,2)"); // 1/2
        OfficeMath eq2 = InsertEquation(builder, @"\r(3,x)"); // cube root of x (unwanted)
        OfficeMath eq3 = InsertEquation(builder, @"\i \su(n=1,5,n)"); // integral with summation

        // Insert a paragraph after the equations.
        builder.Writeln("Paragraph after equations.");

        // Delete the unwanted equation (eq2) and adjust surrounding paragraph spacing.
        DeleteOfficeMath(eq2);

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }

    // Inserts an EQ field, converts it to a real OfficeMath node, and returns the OfficeMath.
    private static OfficeMath InsertEquation(DocumentBuilder builder, string eqCode)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ switch arguments.
        builder.MoveTo(field.Separator);
        builder.Write(eqCode);

        // Update the field so that the EQ code is processed.
        field.Update();

        // Move back to the field start's parent (the paragraph) and start a new paragraph for the next content.
        builder.MoveTo(field.Start.ParentNode);
        builder.InsertParagraph();

        // Convert the field to OfficeMath.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the field.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        field.Remove();

        return officeMath;
    }

    // Removes the specified OfficeMath node and adjusts spacing of surrounding paragraphs.
    private static void DeleteOfficeMath(OfficeMath officeMath)
    {
        if (officeMath == null)
            return;

        // Get the paragraph that contains the OfficeMath.
        Paragraph para = officeMath.ParentParagraph;

        // Remove the OfficeMath node.
        officeMath.Remove();

        // If the paragraph becomes empty after removal, delete it.
        if (!para.HasChildNodes)
        {
            // Keep references to neighboring paragraphs before removing.
            Paragraph prevPara = para.PreviousSibling as Paragraph;
            Paragraph nextPara = para.NextSibling as Paragraph;

            para.Remove();

            // Adjust spacing of neighboring paragraphs if they exist.
            if (prevPara != null)
                prevPara.ParagraphFormat.SpaceAfter = 12; // 12 points after previous paragraph.

            if (nextPara != null)
                nextPara.ParagraphFormat.SpaceBefore = 12; // 12 points before next paragraph.
        }
        else
        {
            // If the paragraph still has content, just adjust its spacing.
            para.ParagraphFormat.SpaceAfter = 12;
            para.ParagraphFormat.SpaceBefore = 12;
        }
    }
}
