using System;
using System.Text;

namespace SalesforceRestAddin.Core;

public static class ExceptionChainText
{
    public static string Format(Exception exception)
    {
        if (exception is null)
        {
            throw new ArgumentNullException(nameof(exception));
        }

        var builder = new StringBuilder();
        AppendChain(builder, exception);
        return builder.ToString().TrimEnd();
    }

    private static void AppendChain(StringBuilder builder, Exception exception)
    {
        var current = exception;
        var depth = 0;
        while (current is not null)
        {
            if (depth > 0)
            {
                builder.AppendLine();
                builder.AppendLine($"--- Inner exception ({depth}) ---");
            }

            builder.AppendLine(current.Message);
            builder.AppendLine($"Type: {current.GetType().FullName}");

            current = current.InnerException;
            depth++;
        }
    }
}
