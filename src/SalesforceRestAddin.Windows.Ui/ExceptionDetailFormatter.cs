using System;
using System.Text;

namespace SalesforceRestAddin.Windows.Ui;

public static class ExceptionDetailFormatter
{
    public static string Format(string? operation, Exception exception)
    {
        if (exception is null)
        {
            throw new ArgumentNullException(nameof(exception));
        }

        var builder = new StringBuilder();
        if (operation is not null && !string.IsNullOrWhiteSpace(operation))
        {
            var trimmed = operation.Trim();
            builder.Append(trimmed);
            if (!trimmed.EndsWith(".", StringComparison.Ordinal))
            {
                builder.Append('.');
            }

            builder.AppendLine();
            builder.AppendLine();
        }

        AppendExceptionChain(builder, exception);
        return builder.ToString().TrimEnd();
    }

    private static void AppendExceptionChain(StringBuilder builder, Exception exception)
    {
        var current = exception;
        var depth = 0;
        while (current is not null)
        {
            if (depth > 0)
            {
                builder.AppendLine();
                builder.AppendLine($"--- Inner exception ({depth}) ---");
                builder.AppendLine();
            }

            builder.AppendLine(current.Message);
            builder.AppendLine($"Type: {current.GetType().FullName}");

            current = current.InnerException;
            depth++;
        }
    }
}
