namespace SalesforceRestAddin.Core.DataPlane;

public static class RecordBatchSplitter
{
    public static IReadOnlyList<IReadOnlyList<T>> Split<T>(IReadOnlyList<T> items, int batchSize)
    {
        if (items is null)
        {
            throw new ArgumentNullException(nameof(items));
        }

        if (batchSize <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(batchSize), "Batch size must be positive.");
        }

        if (items.Count == 0)
        {
            return Array.Empty<IReadOnlyList<T>>();
        }

        var batches = new List<IReadOnlyList<T>>();
        for (var i = 0; i < items.Count; i += batchSize)
        {
            var count = Math.Min(batchSize, items.Count - i);
            var chunk = new T[count];
            for (var j = 0; j < count; j++)
            {
                chunk[j] = items[i + j];
            }

            batches.Add(chunk);
        }

        return batches;
    }
}
