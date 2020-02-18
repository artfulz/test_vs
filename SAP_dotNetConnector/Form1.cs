using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SAP.Middleware.Connector;
using System.Data;

namespace SAP_dotNetConnector
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        public RfcDestination SVIPRD()
        {
            return GetRfcDesctination("BWD", "S_RFC_DGA", "Init!1111", "100", "EN", "10.120.2.168", "01", "10", "600");
            //return GetRfcDesctination("BWP", "S_RFC_DGA", "Dgasap#99", "100", "EN", "10.120.2.171", "01", "10", "600");
        }

        public RfcDestination GetRfcDesctination(string name,
    string username, string password, string client,
    string language, string appServerHost, string systemNumber,
    string maxPollSize, string idleTime)
        {
            RfcConfigParameters parameters = new RfcConfigParameters();
            parameters.Add(RfcConfigParameters.Name, name);                         //Application Name (กำหนดเองได้)
            parameters.Add(RfcConfigParameters.User, username);                     //** Username **
            parameters.Add(RfcConfigParameters.Password, password);                 //** Password **
            parameters.Add(RfcConfigParameters.Client, client);                     //** Client **
            parameters.Add(RfcConfigParameters.Language, language);                 //Language
            parameters.Add(RfcConfigParameters.AppServerHost, appServerHost);       //** Application Server in SAP Connection AppServerHost**
            parameters.Add(RfcConfigParameters.SystemNumber, systemNumber);         //** Instance number in SAP Connection **
            parameters.Add(RfcConfigParameters.PeakConnectionsLimit, maxPollSize);  //PeakConnectionsLimit
            parameters.Add(RfcConfigParameters.IdleTimeout, idleTime);              //IdleTimeout
            return RfcDestinationManager.GetDestination(parameters);
        }

        // class สำหรับ convert sap table    เป็น datatable
        public DataTable GetDataTableFromRFCTable(IRfcTable lrfcTable)
        {
            //sapnco_util
            DataTable loTable = new DataTable();
            //... Create ADO.Net table.
            for (int liElement = 0; liElement < lrfcTable.ElementCount; liElement++)
            {
                RfcElementMetadata metadata = lrfcTable.GetElementMetadata(liElement);
                loTable.Columns.Add(metadata.Name);
            }
            //... Transfer rows from lrfcTable to ADO.Net table.
            foreach (IRfcStructure row in lrfcTable)
            {
                DataRow ldr = loTable.NewRow();
                for (int liElement = 0; liElement < lrfcTable.ElementCount; liElement++)
                {
                    RfcElementMetadata metadata = lrfcTable.GetElementMetadata(liElement);
                    ldr[metadata.Name] = row.GetString(metadata.Name);
                }
                loTable.Rows.Add(ldr);
            }
            return loTable;
        }

        //Class สำหรับ Convert structure เป็น Datatable
        public DataTable ConvertStructure(IRfcStructure myrefcTable)
        {
            DataTable rowTable = new DataTable();
            for (int i = 0; i <= myrefcTable.ElementCount - 1; i++)
            {
                rowTable.Columns.Add(myrefcTable.GetElementMetadata(i).Name);
            }
            DataRow row = rowTable.NewRow();
            for (int j = 0; j <= myrefcTable.ElementCount - 1; j++)
            {
                row[j] = myrefcTable.GetValue(j);
            }
            rowTable.Rows.Add(row);
            return rowTable;
        }





        private void button1_Click_1(object sender, EventArgs e)
        {
            RfcDestination destination = new Form1().SVIPRD();
            IRfcFunction function = destination.Repository.CreateFunction("ZBAPI_PTT_E20_E85_EXPORT");
            //IRfcFunction function = destination.Repository.CreateFunction("ZBAPI_MOE_CUSTOMER");

            //Import Parameters - Value
            function.SetValue("P_FLAG", "X");
            function.SetValue("P_CM_Y_ST", "201901");
            function.SetValue("P_CM_Y_EN", "201912");
            //function.SetValue("CUSTOMER", "");

            //Steucture
            //IRfcStructure struck = function.GetStructure("P_ZOPA_H01");
            //struck.SetValue(0, "X");
            //struck.SetValue(1, "X");
            //struck.SetValue(2, "X");
            //struck.SetValue(3, "X");
            //struck.SetValue(4, "X");
            //struck.SetValue(5, "X");
            //struck.SetValue(6, "X");

            function.Invoke(destination);

            //Selected Tables
            IRfcTable OrderHeader = function.GetTable("ZOS_HXX");
            //IRfcTable OrderHeader = function.GetTable("ZMOE_CUSTOMER");
            DataTable dt = new Form1().GetDataTableFromRFCTable(OrderHeader);
            DataTable dt2 = new Form1().GetDataTableFromRFCTable(OrderHeader);

            int i = dt.Rows.Count;

            dataGridView1.DataSource = dt;
            textBox1.Text = i.ToString();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;
        }
    }
}
