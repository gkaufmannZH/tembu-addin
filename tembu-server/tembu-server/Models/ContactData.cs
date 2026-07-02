namespace tembu_server.Models;

public class ContactData
{
    public string UserEmail { get; set; } = "";
    public string ContactName { get; set; } = "";
    public string ContactEmail { get; set; } = "";
    public List<EmailItem> Emails { get; set; } = [];
    public List<MeetingItem> Meetings { get; set; } = [];
    public List<RumbleItem> Rumbles { get; set; } = [];
}

public class EmailItem
{
    public string DateStr { get; set; } = "";
    public string Direction { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Preview { get; set; } = "";
}

public class MeetingItem
{
    public string DateStr { get; set; } = "";
    public string Subject { get; set; } = "";
    public int DurationMin { get; set; }
}

public class RumbleItem
{
    public string DateStr { get; set; } = "";
    public string Subject { get; set; } = "";
}
