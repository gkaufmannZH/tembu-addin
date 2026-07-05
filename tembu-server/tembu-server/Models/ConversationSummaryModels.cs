namespace tembu_server.Models;

public class ConversationSummaryRequest
{
    public string Lang { get; set; } = "";
    public string ContactName { get; set; } = "";
    public string ContactEmail { get; set; } = "";
    public List<ConversationInput> Conversations { get; set; } = [];
}

public class ConversationInput
{
    // Vom Client vergeben (Graph conversationId) — wird 1:1 in der Antwort zurückgegeben,
    // damit der Client die Zusammenfassung der richtigen Unterhaltung zuordnen kann.
    public string Id { get; set; } = "";
    public string Subject { get; set; } = "";
    public List<EmailItem> Emails { get; set; } = [];
}
