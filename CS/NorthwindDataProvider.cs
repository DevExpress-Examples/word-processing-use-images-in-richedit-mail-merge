using System.Data;
using System.Data.OleDb;

namespace RichEditImageMailMerge {
    public static class NorthwindDataProvider {
        private static string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\nwind.mdb;Persist Security Info=False;";

        static NorthwindDataProvider() {
            using (OleDbConnection connection = new OleDbConnection(connectionString)) {
                OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Categories", connection);
                OleDbDataAdapter da = new OleDbDataAdapter(selectCommand);

                categories = new DataTable("Categories");

                da.Fill(categories);

                categories.Constraints.Add("IDPK", categories.Columns["CategoryID"], true);
 
                selectCommand.Dispose();
            }
        }

        private static DataTable categories;

        public static DataTable Categories {
            get {
                return categories;
            }
        }
    }
}