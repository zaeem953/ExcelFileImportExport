namespace adodotnet.Utility
{
	public static class ConnectionString
	{
		private static string _connectionString= "server=localhost;user=root;database=exceldata;password=1234;port=3306;AllowLoadLocalInFile=true";
		public static string  _ConnectionString { get => _connectionString;  }
	}
}
