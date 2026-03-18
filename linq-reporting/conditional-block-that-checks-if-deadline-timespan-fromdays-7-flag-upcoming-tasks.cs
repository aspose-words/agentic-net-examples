using System;
using Aspose.Words;
using Aspose.Words.Fields;

class UpcomingTaskFlag
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a label for the task.
        builder.Writeln("Task:");

        // Insert an IF field that evaluates whether the merge field "deadline"
        // (expected to contain the number of days until the deadline) is less than 7.
        // If true, the field will display the word "Upcoming".
        FieldIf ifField = (FieldIf)builder.InsertField(FieldType.FieldIf, true);
        // Left side of the comparison – a MERGEFIELD named "deadline".
        ifField.LeftExpression = "MERGEFIELD deadline";
        // Comparison operator.
        ifField.ComparisonOperator = "<";
        // Right side – the constant value 7 (days).
        ifField.RightExpression = "7";
        // Text to display when the condition is true.
        ifField.TrueText = "Upcoming";
        // Text to display when the condition is false (empty in this case).
        ifField.FalseText = "";
        // Evaluate the field so the result appears in the document.
        ifField.Update();

        // Save the document.
        doc.Save("UpcomingTasks.docx");
    }
}
