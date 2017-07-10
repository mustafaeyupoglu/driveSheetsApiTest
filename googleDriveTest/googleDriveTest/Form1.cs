using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Threading;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
namespace googleDriveTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "testsheete2";

        private void button1_Click(object sender, EventArgs e)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("Content/client_secret.json", FileMode.Open, FileAccess.Read))
            {  string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/drive-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            String spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
            String range = "A1:K";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
            ValueRange response = request.Execute();

            IList<IList<Object>> values = response.Values;
            List<temp> liste = new List<temp>();
            if (values != null && values.Count > 0)
            {
                

                foreach (var row in values)
                {
                    temp t = new temp { id1 = row[0].ToString(), id2 = row[1].ToString() };
                    liste.Add(t);

                }
            }
            else
            {
                MessageBox.Show("No data found.");
            }
            List<temp> tt = liste.OrderBy(x => x.id2).ToList<temp>();
         
            dataGridView1.DataSource = tt;
            dataGridView1.Refresh();
        }

        class temp
        {
           public string id1 { get; set; }
           public string id2 { get; set; }
           public string id3 { get; set; }
           public string id4 { get; set; }
           public string id5 { get; set; }


        }
    }
}
