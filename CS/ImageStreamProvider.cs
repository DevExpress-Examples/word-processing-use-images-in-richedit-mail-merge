using System.IO;
using System.Data;
using DevExpress.Office.Services;

namespace RichEditImageMailMerge {
#region #iuristreamprovider
    public class ImageStreamProvider : IUriStreamProvider {
        static readonly string prefix = "dbimg://";
        DataTable table;
        string columnName;

        public ImageStreamProvider(DataTable sourceTable, string imageColumn) {
            this.table = sourceTable;
            this.columnName = imageColumn;
        }


        public Stream GetStream(string uri) {
            uri = uri.Trim();
            if (!uri.StartsWith(prefix))
                return null;
            string strId = uri.Substring(prefix.Length).Trim();
            int id;
            if (!int.TryParse(strId, out id))
                return null;
            DataRow row = table.Rows.Find(id);
            if (row == null)
                return null;
            byte[] bytes = row[columnName] as byte[];
            if (bytes == null)
                return null;

            MemoryStream memoryStream = new MemoryStream(bytes);
            return memoryStream;
        }

#endregion #iuristreamprovider
    }
}