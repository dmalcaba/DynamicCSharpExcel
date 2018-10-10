namespace DynamicCsharp
{
    /// <summary>
    /// For Test Data
    /// </summary>
    public class AccountValues
    {
        public string AccountCode { get; set; }
        public string AccountName { get; set; }
        public string CostCenter { get; set; }
        public decimal Amount { get; set; }

        public AccountValues(string accountCode, string accountName, string costCenter, decimal amount)
        {
            AccountCode = accountCode;
            AccountName = accountName;
            CostCenter = costCenter;
            Amount = amount;
        }
    }
}
