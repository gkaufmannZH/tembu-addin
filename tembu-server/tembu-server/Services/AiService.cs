using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Options;
using tembu_server.Models;

namespace tembu_server.Services;

public class AiService
{
    private readonly HttpClient _http = new();
    private readonly IOptionsMonitor<AiSettings> _settings;

    public AiService(IOptionsMonitor<AiSettings> settings)
    {
        _settings = settings;
    }

    public Task<string> CallAsync(string prompt)
    {
        var s = _settings.CurrentValue;
        return s.Provider switch
        {
            "openai"    => CallOpenAICompatAsync(prompt, s.ApiKey, "https://api.openai.com/v1", s.Model.Length > 0 ? s.Model : "gpt-4o-mini"),
            "anthropic" => CallAnthropicAsync(prompt, s),
            "groq"      => CallOpenAICompatAsync(prompt, s.ApiKey, "https://api.groq.com/openai/v1", s.Model.Length > 0 ? s.Model : "llama-3.1-70b-versatile"),
            "ollama"    => CallOpenAICompatAsync(prompt, null, (s.Endpoint.Length > 0 ? s.Endpoint : "http://localhost:11434") + "/v1", s.Model.Length > 0 ? s.Model : "qwen2.5:14b"),
            _           => CallGeminiAsync(prompt, s)
        };
    }

    private async Task<string> CallGeminiAsync(string prompt, AiSettings s)
    {
        var mdl = s.Model.Length > 0 ? s.Model : "gemini-2.5-flash";
        var url = $"https://generativelanguage.googleapis.com/v1beta/models/{mdl}:generateContent?key={s.ApiKey}";
        var body = JsonSerializer.Serialize(new
        {
            contents = new[] { new { parts = new[] { new { text = prompt } } } },
            generationConfig = new { temperature = 0.3, maxOutputTokens = 8192 }
        });
        var response = await _http.PostAsync(url, new StringContent(body, Encoding.UTF8, "application/json"));
        var json = await response.Content.ReadAsStringAsync();
        var doc  = TryParse(json);
        if (!response.IsSuccessStatusCode || doc == null || !doc.RootElement.TryGetProperty("candidates", out var candidates))
            throw ProviderError("Gemini", response, json, doc);
        return candidates[0].GetProperty("content").GetProperty("parts")[0].GetProperty("text").GetString() ?? "";
    }

    private async Task<string> CallAnthropicAsync(string prompt, AiSettings s)
    {
        var mdl = s.Model.Length > 0 ? s.Model : "claude-haiku-4-5-20251001";
        using var req = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages");
        req.Headers.Add("x-api-key", s.ApiKey);
        req.Headers.Add("anthropic-version", "2023-06-01");
        req.Content = new StringContent(JsonSerializer.Serialize(new
        {
            model = mdl, max_tokens = 4096,
            messages = new[] { new { role = "user", content = prompt } }
        }), Encoding.UTF8, "application/json");
        var response = await _http.SendAsync(req);
        var json = await response.Content.ReadAsStringAsync();
        var doc  = TryParse(json);
        if (!response.IsSuccessStatusCode || doc == null || !doc.RootElement.TryGetProperty("content", out var content))
            throw ProviderError("Anthropic", response, json, doc);
        return content[0].GetProperty("text").GetString() ?? "";
    }

    private async Task<string> CallOpenAICompatAsync(string prompt, string? apiKey, string baseUrl, string model)
    {
        using var req = new HttpRequestMessage(HttpMethod.Post, $"{baseUrl}/chat/completions");
        if (apiKey != null) req.Headers.Add("Authorization", $"Bearer {apiKey}");
        req.Content = new StringContent(JsonSerializer.Serialize(new
        {
            model,
            messages  = new[] { new { role = "user", content = prompt } },
            max_tokens  = 4096,
            temperature = 0.3
        }), Encoding.UTF8, "application/json");
        var response = await _http.SendAsync(req);
        var json = await response.Content.ReadAsStringAsync();
        var doc  = TryParse(json);
        if (!response.IsSuccessStatusCode || doc == null || !doc.RootElement.TryGetProperty("choices", out var choices))
            throw ProviderError(model, response, json, doc);
        return choices[0].GetProperty("message").GetProperty("content").GetString() ?? "";
    }

    // ── Fehlerbehandlung ──────────────────────────────────────────────────
    // Provider-Antworten bei Fehlern haben nicht die erwartete Erfolgs-Form
    // (z.B. {"error":{"message":"..."}} statt {"candidates":[...]}) — ohne
    // diese Prüfung wirft GetProperty() eine nichtssagende KeyNotFoundException.
    private static JsonDocument? TryParse(string json)
    {
        try { return JsonDocument.Parse(json); } catch (JsonException) { return null; }
    }

    private static Exception ProviderError(string provider, HttpResponseMessage response, string raw, JsonDocument? doc)
    {
        string detail = raw;
        if (doc != null && doc.RootElement.TryGetProperty("error", out var err))
        {
            if (err.ValueKind == JsonValueKind.Object && err.TryGetProperty("message", out var msg))
                detail = msg.GetString() ?? raw;
            else if (err.ValueKind == JsonValueKind.String)
                detail = err.GetString() ?? raw;
        }
        return new Exception($"{provider} ({(int)response.StatusCode}): {detail.Trim()}");
    }
}
