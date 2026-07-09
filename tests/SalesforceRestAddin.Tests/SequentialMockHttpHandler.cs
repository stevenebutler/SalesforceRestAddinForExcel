using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SalesforceRestAddin.Tests;

/// <summary>
/// Returns canned responses in request order, routed by URL path for concurrent-safe choreography.
/// </summary>
public sealed class SequentialMockHttpHandler : HttpMessageHandler
{
    private readonly Dictionary<string, Queue<Func<HttpRequestMessage, HttpResponseMessage>>> _routes = new(
        StringComparer.OrdinalIgnoreCase);
    private readonly List<HttpRequestMessage> _requests = new();

    public IReadOnlyList<HttpRequestMessage> Requests => _requests;

    public void Enqueue(HttpStatusCode statusCode, string body, string contentType = "application/json") =>
        EnqueueForRoute("default", statusCode, body, contentType);

    public void Enqueue(Func<HttpRequestMessage, HttpResponseMessage> responder) =>
        EnqueueForRoute("default", responder);

    public void EnqueueForTokenEndpoint(HttpStatusCode statusCode, string body) =>
        EnqueueForRoute("oauth2/token", statusCode, body);

    public void EnqueueForIdentity(HttpStatusCode statusCode, string body) =>
        EnqueueForRoute("/id/", statusCode, body);

    public void EnqueueForRoute(
        string pathContains,
        HttpStatusCode statusCode,
        string body,
        string contentType = "application/json") =>
        EnqueueForRoute(pathContains, _ => new HttpResponseMessage(statusCode)
        {
            Content = new StringContent(body, Encoding.UTF8, contentType),
        });

    public void EnqueueForRoute(string pathContains, Func<HttpRequestMessage, HttpResponseMessage> responder)
    {
        if (!_routes.TryGetValue(pathContains, out var queue))
        {
            queue = new Queue<Func<HttpRequestMessage, HttpResponseMessage>>();
            _routes[pathContains] = queue;
        }

        queue.Enqueue(responder);
    }

    protected override Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request,
        CancellationToken cancellationToken)
    {
        _requests.Add(request);

        var path = request.RequestUri?.AbsolutePath ?? string.Empty;
        var routeKey = ResolveRouteKey(path);
        if (!_routes.TryGetValue(routeKey, out var queue) || queue.Count == 0)
        {
            throw new InvalidOperationException(
                $"No mock response configured for {request.Method} {request.RequestUri} (route '{routeKey}')");
        }

        var responder = queue.Dequeue();
        return Task.FromResult(responder(request));
    }

    private static string ResolveRouteKey(string path)
    {
        if (path.Contains("oauth2/token", StringComparison.OrdinalIgnoreCase))
        {
            return "oauth2/token";
        }

        if (path.Contains("/id/", StringComparison.OrdinalIgnoreCase))
        {
            return "/id/";
        }

        return "default";
    }
}
