using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace SalesforceRestAddin;

/// <summary>
/// Small async FIFO used to bridge background REST work to Excel STA work without
/// buffering an unbounded number of query pages. This has one producer and one consumer.
/// </summary>
internal sealed class BoundedAsyncQueue<T>
{
    private readonly object _gate = new();
    private readonly Queue<T> _items = new();
    private readonly SemaphoreSlim _availableSlots;
    private readonly SemaphoreSlim _availableItems = new(0);
    private bool _completed;

    public BoundedAsyncQueue(int capacity)
    {
        if (capacity <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(capacity));
        }

        _availableSlots = new SemaphoreSlim(capacity, capacity);
    }

    public async Task EnqueueAsync(T item, CancellationToken cancellationToken)
    {
        await _availableSlots.WaitAsync(cancellationToken).ConfigureAwait(false);
        lock (_gate)
        {
            if (_completed)
            {
                _availableSlots.Release();
                throw new InvalidOperationException("Cannot enqueue after completion.");
            }

            _items.Enqueue(item);
        }

        _availableItems.Release();
    }

    public async Task<(bool HasItem, T? Item)> TryDequeueAsync()
    {
        while (true)
        {
            await _availableItems.WaitAsync().ConfigureAwait(false);
            lock (_gate)
            {
                if (_items.Count > 0)
                {
                    var item = _items.Dequeue();
                    _availableSlots.Release();
                    return (true, item);
                }

                if (_completed)
                {
                    return (false, default);
                }
            }
        }
    }

    public void Complete()
    {
        lock (_gate)
        {
            if (_completed)
            {
                return;
            }

            _completed = true;
        }

        // There is a single consumer. Wake it if it is waiting for the terminal state.
        _availableItems.Release();
    }
}
