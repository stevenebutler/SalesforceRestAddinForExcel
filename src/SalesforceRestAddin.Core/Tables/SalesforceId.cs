namespace SalesforceRestAddin.Core.Tables;

public static class SalesforceId
{
    private const string SuffixChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ012345";

    public static string? Normalize(string? id)
    {
        if (id is null || string.IsNullOrWhiteSpace(id))
        {
            return null;
        }

        id = id.Trim();
        if (id.Length == 18)
        {
            return id;
        }

        if (id.Length != 15)
        {
            return id;
        }

        var suffix = new char[3];
        for (var i = 0; i < 3; i++)
        {
            var flags = 0;
            for (var j = 0; j < 5; j++)
            {
                if (char.IsUpper(id[i * 5 + j]))
                {
                    flags |= 1 << j;
                }
            }

            suffix[i] = SuffixChars[flags];
        }

        return id + new string(suffix);
    }

    public static bool IsValid(string? id) =>
        id is not null
        && !string.IsNullOrWhiteSpace(id)
        && (id.Trim().Length == 15 || id.Trim().Length == 18);
}
