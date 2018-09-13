using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.SqlClient;
using System.Diagnostics;
using InsertMO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;


namespace GP_MOGen_Verification
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public List<Site> source = new List<Site>();

        string GPconString =
        "SERVER=RGV-HP-SQL01;" +
        "DATABASE=UC;" +
        "USER=sa;" +
        "PASSWORD=Pass@word1;";

        public MainWindow()
        {
            InitializeComponent();
            source.Add(new GP_MOGen_Verification.Site() { SiteID = "UC" });
            source.Add(new GP_MOGen_Verification.Site() { SiteID = "WAREHOUSE" });
            cmb.ItemsSource = source;
            this.DataContext = this;
        }

        private void VerifyButton_Click(object sender, RoutedEventArgs e)
        {
            string SITE = "";
            if (cmb.SelectedIndex != -1)
            {
                var LOCN = (Site)cmb.SelectedItem;
                SITE = LOCN.SiteID;

            }
            //first do data checks to ensure no fields are left blank
            if (Model_Textbox.Text != "" && SN_Textbox.Text != "" && SITE != "")
            {
                //data is in all fields, now verify ITEMNMBR
                bool itemcheck = false;

                itemcheck = ITEMNMBRCHECK(Model_Textbox.Text.Trim());

                if (itemcheck == true)
                {
                    string SN = SN_Textbox.Text;
                    string ITEM = "";
                    string LOCNCODE = "";
                    DateTime date1 = new DateTime();
                    int status = 0;
                    string MONumber = "";
                    int QTY = 0;

                    if (SerializedCheckbox.IsChecked == true)
                    {
                        // Look up SN in W7158MoHdr Table
                        using (SqlConnection con = new SqlConnection(GPconString))
                        {
                            using (SqlCommand cmd = con.CreateCommand())
                            {

                                cmd.CommandText = "SELECT ITEMNMBR, LOCNCODE, DATE1, STATUS, MANUFACTUREORDER_I"
                                    + " FROM W7158MoHdr"
                                    + " WHERE SERLTNUM = @SN";

                                cmd.Parameters.AddWithValue("@SN", SN);

                                con.Open();
                                SqlDataReader reader = cmd.ExecuteReader();
                                while (reader.Read())
                                {
                                    ITEM = reader.GetString(0).Trim();
                                    LOCNCODE = reader.GetString(1).Trim();
                                    date1 = reader.GetDateTime(2);
                                    status = reader.GetInt16(3);
                                    MONumber = reader.GetString(4).Trim();
                                }
                            }
                        }
                        DisplayMessage(MONumber, ITEM, 1, SN, LOCNCODE, date1, status);
                    }
                    else
                    {
                        // Look up LAST MO in W7158MoHdr Table
                        using (SqlConnection con = new SqlConnection(GPconString))
                        {
                            using (SqlCommand cmd = con.CreateCommand())
                            {

                                cmd.CommandText = "SELECT TOP 1 MANUFACTUREORDER_I, ITEMNMBR, QUANTITY,LOCNCODE, DATE1, STATUS"
                                    + " FROM W7158MoHdr"
                                    + " WHERE ITEMNMBR = @ITEM"
                                    + " ORDER BY MANUFACTUREORDER_I DESC";

                                cmd.Parameters.AddWithValue("@ITEM", Model_Textbox.Text.Trim());

                                con.Open();
                                SqlDataReader reader = cmd.ExecuteReader();
                                while (reader.Read())
                                {
                                    MONumber = reader.GetString(0).Trim();
                                    ITEM = reader.GetString(1).Trim();
                                    QTY = reader.GetInt32(2);
                                    LOCNCODE = reader.GetString(3).Trim();
                                    date1 = reader.GetDateTime(4);
                                    status = reader.GetInt16(5);
                                }
                            }
                        }

                        DisplayMessage(MONumber, ITEM, QTY, "", LOCNCODE, date1, status);
                    }
                }
                else
                {
                    MessageBox.Show("Entered invalid item number -- must enter valid item number");
                    ClearAllValues();
                }
            }
        }

        private void ClearAllValues()
        {
            Model_Textbox.Clear();
            SN_Textbox.Clear();
            cmb.SelectedIndex = -1;
            qtyTextbox.Clear();

        }

        private bool ITEMNMBRCHECK(string ITEMNMBR)
        {
            bool isValid = false;
            string DESCRIPTION = "";

            try
            {

                Debug.WriteLine("IN ITEMNMBRCHECK ITEMNMBR = " + ITEMNMBR);
                // need to get ITEMDESC
                using (SqlConnection con = new SqlConnection(GPconString))
                {
                    using (SqlCommand cmd = con.CreateCommand())
                    {

                        cmd.CommandText = "SELECT ITEMDESC"
                            + " FROM IV00101"
                            + " WHERE ITEMNMBR = @ITEMNMBR";

                        cmd.Parameters.AddWithValue("@ITEMNMBR", ITEMNMBR);

                        con.Open();
                        SqlDataReader reader = cmd.ExecuteReader();
                        while (reader.Read())
                        {
                            DESCRIPTION = reader.GetString(0).Trim();
                        }
                    }
                }

                if (DESCRIPTION != "")
                {
                    isValid = true;
                    return isValid;
                }
                else
                {
                    isValid = false;
                    return isValid;
                }
            }
            catch
            {
                MessageBox.Show("Errored during Item Number Check");
                isValid = false;
                return isValid;
            }
        }

        private void InsertButton_Click(object sender, RoutedEventArgs e)
        {
            string app = "GP MOGen Verification";
            string user = Environment.UserName;

            string SITE = "";

            if (cmb.SelectedIndex != -1)
            {
                var LOCN = (Site)cmb.SelectedItem;
                SITE = LOCN.SiteID;
            }

            Debug.WriteLine("COMBOBOX TEXT = " + SITE);

            //first do data checks to ensure no fields are left blank
            if (Model_Textbox.Text != "" && SN_Textbox.Text != "" && SITE != "")
            {
                if (SerializedCheckbox.IsChecked == true)
                {
                    CreateMO newMO = new CreateMO();
                    string message = newMO.InsertMOwithSN(SN_Textbox.Text.Trim(), Model_Textbox.Text.Trim(), 1, SITE, app, user);
                    MessageBox.Show(message);
                }
            }
            else if (Model_Textbox.Text == "")
            {
                MessageBox.Show("Item Number field blank! Must have item number");
            }
            else if (SITE == "")
            {
                MessageBox.Show("No site selected");
            //else if (SN_Textbox.Text == "")
            //{
            //    MessageBox.Show("Serial Number field blank! If Serialized is checked, must have serial number");
            }
            else if (SN_Textbox.Text == "" && SerializedCheckbox.IsChecked == true)
            {
                MessageBox.Show("Serial Number field blank! If Serialized is checked, must have serial number");
            }
            else if (SN_Textbox.Text == "" && SerializedCheckbox.IsChecked == false)
            {
                CreateMO newMO = new CreateMO();
                string message = newMO.InsertMOwithoutSN(Model_Textbox.Text.Trim(), Convert.ToInt32(qtyTextbox.Text), SITE, app, user);
                MessageBox.Show(message);
            }

            ClearAllValues();
        }

        private void CheckBoxUnchecked(object sender, RoutedEventArgs e)
        {
            qtyTextbox.IsEnabled = true;
            SN_Textbox.IsEnabled = false;
        }

        private void CheckBoxChecked(object sender, RoutedEventArgs e)
        {
            qtyTextbox.IsEnabled = false;
            SN_Textbox.IsEnabled = true;
        }

        private void ReportButton_Click(object sender, RoutedEventArgs e)
        {
            //first do data checks to ensure ITEM fields are not left blank
            if (Model_Textbox.Text != "")
            {

                //create excel sheet
                Excel.Application oXL;
                Excel._Workbook oWB;
                Excel._Worksheet oSheet;

                try
                {
                    //Start Excel and get Application object.
                    oXL = new Excel.Application();
                    oXL.Visible = true;

                    //Get a new workbook.
                    oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                    //TO RENAME THE CREATED EXCEL FILE 
                    //object missing = System.Reflection.Missing.Value;

                    oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                    //Name the sheet.
                    oSheet.Name = "MOGen MO Report";
                    //Add table headers going cell by cell.
                    oWB.ActiveSheet.Cells[1, 1] = "MO Number";
                    oWB.ActiveSheet.Cells[1, 1].Font.Bold = true;
                    oWB.ActiveSheet.Cells[1, 2] = "Item Number";
                    oWB.ActiveSheet.Cells[1, 2].Font.Bold = true;
                    oWB.ActiveSheet.Cells[1, 3] = "Quantity";
                    oWB.ActiveSheet.Cells[1, 3].Font.Bold = true;
                    oWB.ActiveSheet.Cells[1, 4] = "Serial Number";
                    oWB.ActiveSheet.Cells[1, 4].Font.Bold = true;
                    oWB.ActiveSheet.Cells[1, 5] = "Date";
                    oWB.ActiveSheet.Cells[1, 5].Font.Bold = true;
                    oWB.ActiveSheet.Cells[1, 6] = "Site";
                    oWB.ActiveSheet.Cells[1, 6].Font.Bold = true;
                    oWB.ActiveSheet.Cells[1, 7] = "Status";
                    oWB.ActiveSheet.Cells[1, 7].Font.Bold = true;


                    string gpQuery =
                        "SELECT MANUFACTUREORDER_I, ITEMNMBR, QUANTITY, SERLTNUM, DATE1, LOCNCODE, STATUS"
                        + " FROM W7158MOHdr WHERE ITEMNMBR = @ITEM"
                        + " ORDER BY MANUFACTUREORDER_I DESC;";

                    Debug.WriteLine("Query = " + gpQuery);

                    string MONumber = "";
                    string ITEMNMBR = "";
                    decimal QUANTITY = 0m;
                    string SERLTNUM = "";
                    DateTime date1 = new DateTime();
                    string SITE = "";
                    Int32 status = 0;
                    string stringStatus = "";


                    // Create and open the connection in a using block. This
                    // ensures that all resources will be closed and disposed
                    // when the code exits.
                    using (SqlConnection connection = new SqlConnection(GPconString))
                    {
                        // Create the Command and Parameter objects.
                        SqlCommand command = new SqlCommand(gpQuery, connection);
                        command.Parameters.AddWithValue("@ITEM", Model_Textbox.Text);

                        try
                        {
                            int currentrow = 1;
                            connection.Open();

                            SqlDataReader reader = command.ExecuteReader();
                            while (reader.Read())
                            {
                                currentrow = currentrow + 1;
                                MONumber = reader.GetString(0).Trim();
                                ITEMNMBR = reader.GetString(1).Trim();
                                QUANTITY = reader.GetDecimal(2);
                                SERLTNUM = reader.GetString(3).Trim();
                                date1 = reader.GetDateTime(4);
                                SITE = reader.GetString(5).Trim();
                                status = Convert.ToInt32(reader.GetInt16(6));


                                switch (status)
                                {
                                    case 2:
                                        stringStatus = "Completed";
                                        break;

                                    case 1:
                                        stringStatus = "Failed";
                                        break;

                                    case 9999:
                                        stringStatus = "In Process";
                                        break;

                                    case 0:
                                        stringStatus = "Unprocessed";
                                        break;
                                }

                                oWB.ActiveSheet.Cells[currentrow, 1] = MONumber;
                                oWB.ActiveSheet.Cells[currentrow, 2] = ITEMNMBR;
                                oWB.ActiveSheet.Cells[currentrow, 3] = QUANTITY.ToString();
                                oWB.ActiveSheet.Cells[currentrow, 4] = SERLTNUM;
                                oWB.ActiveSheet.Cells[currentrow, 5] = date1.ToString("yyyyMMdd");
                                oWB.ActiveSheet.Cells[currentrow, 6] = SITE;
                                oWB.ActiveSheet.Cells[currentrow, 7] = stringStatus;
                            }
                        }
                        catch
                        {
                            Debug.WriteLine("Caught exception writing to excel");
                        }
                    }
                }
                catch
                {
                    Debug.WriteLine("Caught Exception opening excel");
                }

                MessageBox.Show("Report is finished");

            }
        }

        private void DisplayMessage(string MONumber, string ITEM, int QTY, string SN, string LOCNCODE, DateTime date1, int status)
        {
            string message = "";

            switch (status)
            {
                case 2:
                    message = "Complete";
                    break;
                case 1:
                    message = "Failed";
                    break;
                case 9999:
                    message = "In Process";
                    break;
                case 0:
                    message = "Unprocessed";
                    break;
            }
            MessageBox.Show("MO Number:" + MONumber + Environment.NewLine
            + "Item Number:" + ITEM.Trim() + Environment.NewLine
            + "Serial Number:" + SN.Trim() + Environment.NewLine
            + "Quantity:" + QTY.ToString() + Environment.NewLine
            + "Site:" + LOCNCODE.Trim() + Environment.NewLine
            + "Date:" + date1.ToString("yyyy-MM-dd") + Environment.NewLine
            + "Status:" + message);
            ClearAllValues();
        }
        

        }

    }


