namespace tembu_server.Models;

public class AiSettings
{
    public string Provider { get; set; } = "gemini";
    public string ApiKey   { get; set; } = "";
    public string Model    { get; set; } = "";
    public string Endpoint { get; set; } = "";
}
