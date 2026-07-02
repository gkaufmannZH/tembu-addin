namespace tembu_server.Models;

public class CompanyData
{
    public string UserEmail   { get; set; } = "";
    public string CompanyName { get; set; } = "";
    public string Domain      { get; set; } = "";
    public List<CompanyContact> Contacts { get; set; } = [];
    public List<CompanyEmailItem> Emails { get; set; } = [];
    public List<CompanyMeetingItem> Meetings { get; set; } = [];
}

public class CompanyContact
{
    public string Name  { get; set; } = "";
    public string Email { get; set; } = "";
}

public class CompanyEmailItem
{
    public string DateStr   { get; set; } = "";
    public string Direction { get; set; } = "";
    public string Contact   { get; set; } = "";
    public string Subject   { get; set; } = "";
}

public class CompanyMeetingItem
{
    public string DateStr { get; set; } = "";
    public string Contact { get; set; } = "";
    public string Subject { get; set; } = "";
    public int DurationMin { get; set; }
}
