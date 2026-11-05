using System;
using System.IO;

public class Program
{
    // Path to the audit log file.
    private const string AuditFilePath = "audit.log";

    // Entry point of the console application.
    public static void Main()
    {
        // Simulate a cancellation event.
        LogCancellationEvent("User requested cancellation of the operation.");

        // Simulate another cancellation event.
        LogCancellationEvent("System timeout triggered cancellation.");

        // Verify that the audit file was created.
        if (!File.Exists(AuditFilePath))
        {
            throw new InvalidOperationException($"Audit file was not created at path: {AuditFilePath}");
        }

        // Optionally, display the contents of the audit file (not required for compliance).
        // Console.WriteLine(File.ReadAllText(AuditFilePath));
    }

    // Logs a cancellation event with a timestamp to the audit file.
    private static void LogCancellationEvent(string description)
    {
        string timestamp = DateTime.UtcNow.ToString("o"); // ISO 8601 format.
        string logEntry = $"{timestamp} - Cancellation: {description}{Environment.NewLine}";

        // Append the log entry to the audit file.
        File.AppendAllText(AuditFilePath, logEntry);
    }
}
