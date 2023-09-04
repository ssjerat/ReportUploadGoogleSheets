public class DataReport
{
    public string clientID { get; init; }
    public string clientName { get; init; }
    public string clientAmount { get; set; }

    public static List<List<DataReport>> ListDataReports = new List<List<DataReport>>();
    public DataReport(string clientID, string clientName, string clientAmount)
    {
        this.clientID = clientID;
        this.clientName = clientName;
        this.clientAmount = clientAmount;
    }
    public override int GetHashCode()
    {
        return HashCode.Combine(clientID, clientName);
    }

    public override bool Equals(object? obj)
    {
        if (obj == null || GetType() != obj.GetType())
            return false;

        DataReport other = (DataReport)obj;
        return clientID == other.clientID;
    }
}
