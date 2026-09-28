using System;
using System.IO;

public class Program
{
    // Path to the audit log file.
    private const string AuditFilePath = "audit.log";

    public static void Main()
    {
        // Ensure the audit file exists; create if it does not.
        if (!File.Exists(AuditFilePath))
        {
            using (File.Create(AuditFilePath)) { }
        }

        // Log a cancellation event with the current UTC timestamp.
        LogCancellationEvent();

        // Validate that the log entry was written.
        ValidateLogEntry();

        // Indicate successful completion (optional, not required for compliance).
        // Console.WriteLine("Cancellation event logged successfully.");
    }

    private static void LogCancellationEvent()
    {
        string timestamp = DateTime.UtcNow.ToString("o"); // ISO 8601 format.
        string logEntry = $"{timestamp} - Cancellation event recorded.";

        // Append the log entry to the audit file.
        File.AppendAllText(AuditFilePath, logEntry + Environment.NewLine);
    }

    private static void ValidateLogEntry()
    {
        // Read the last line from the audit file.
        string[] lines = File.ReadAllLines(AuditFilePath);
        if (lines.Length == 0)
        {
            throw new InvalidOperationException("Audit file is empty after logging.");
        }

        string lastLine = lines[^1]; // C# 8.0 index from end.
        if (!lastLine.Contains("Cancellation event recorded"))
        {
            throw new InvalidOperationException("The last audit entry does not contain the expected cancellation message.");
        }
    }
}
