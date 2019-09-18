using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace exportarDGVaEXCEL
{
    public partial class Form1 : Form
    {
        Conexion con = new Conexion();
        public Form1()
        {
            InitializeComponent();
        }
        int i = 0;
        private void btnExportar_Click(object sender, EventArgs e)
        {
            ExportarDataGridViewExcel(dgvResultado);
        }
        private void ExportarDataGridViewExcel(DataGridView grd)
        {
            var csv = new StringBuilder();
            foreach (DataGridViewRow row in dgvResultado.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }
                var text = "";
                foreach (DataGridViewCell cell in row.Cells)
                {
                    text += cell.Value.ToString() ?? "" + ",";
                }
                text.Remove(text.Length - 1, 1);
                text += "\n";
                csv.AppendLine(text);
            }
            string ruta = "c:\\DGV-CSV\\";
            if (!Directory.Exists(ruta))
            {
                Directory.CreateDirectory(ruta);
            }
            File.WriteAllText(ruta + i + ".csv", csv.ToString());
            /*SaveFileDialog fichero = new SaveFileDialog();
            fichero.Filter = "Excel (*.xls)|*.xls";
            if (fichero.ShowDialog() == DialogResult.OK)
            {
                Microsoft.Office.Interop.Excel.Application aplicacion;
                Microsoft.Office.Interop.Excel.Workbook libros_trabajo;
                Microsoft.Office.Interop.Excel.Worksheet hoja_trabajo;
                aplicacion = new Microsoft.Office.Interop.Excel.Application();
                libros_trabajo = aplicacion.Workbooks.Add();
                hoja_trabajo =
                    (Microsoft.Office.Interop.Excel.Worksheet)libros_trabajo.Worksheets.get_Item(1);
                //Recorremos el DataGridView rellenando la hoja de trabajo
                for (int i = 0; i < grd.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < grd.Columns.Count; j++)
                    {
                        hoja_trabajo.Cells[i + 1, j + 1] = grd.Rows[i].Cells[j].Value.ToString();
                    }
                }
                libros_trabajo.SaveAs(fichero.FileName,
                    Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);
                libros_trabajo.Close(true);
                aplicacion.Quit();
            }*/
        }

        private void BtnConsultar_Click(object sender, EventArgs e)
        {
            dgvResultado.Rows.Clear();
            dgvResultado.Columns.Clear();
            string conectionString;


            string query = generarConsulta();
            MySqlConnection databaseConnection = con.ObtenerMySQLConnection();
            MySqlCommand commandDatabase = new MySqlCommand(query, databaseConnection);
            commandDatabase.CommandTimeout = 60;
            MySqlDataReader reader;
            try
            {
                // Abre la base de datos
                if (databaseConnection.State != ConnectionState.Open) { databaseConnection.Open(); }

                // Ejecuta la consultas
                reader = commandDatabase.ExecuteReader();
                // Si tu consulta retorna un resultado, guarda los datos en la tabla dgvListado'
                if (reader.HasRows)
                {
                    DataGridViewRow nombres = new DataGridViewRow();
                    if (dgvResultado.Columns.Count < 1)
                    {
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            DataGridViewCell celda = new DataGridViewTextBoxCell();
                            celda.Value = reader.GetName(i).ToString();//Set Cell Value
                            nombres.Cells.Add(celda);
                            dgvResultado.Columns.Add(reader.GetName(i).ToString(), reader.GetName(i).ToString());

                            if (i == reader.FieldCount - 1 && i < 9)
                            {
                                dgvResultado.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                            }
                        }
                    }
                    while (reader.Read())
                    {

                        DataGridViewRow valores = new DataGridViewRow();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            DataGridViewCell celda = new DataGridViewTextBoxCell();

                            string val;
                            try
                            {
                                val = (reader.GetString(nombres.Cells[i].Value.ToString())) ?? "";

                            }
                            catch (Exception)
                            {
                                val = "";
                            }
                            celda.Value = val;
                            valores.Cells.Add(celda);
                        }
                        dgvResultado.Rows.Add(valores);
                    }

                    databaseConnection.Close();
                }
                else
                {
                    MessageBox.Show("No se encontraron registros.");
                }
            }
            catch (Exception ex) { MessageBox.Show("Error:" + ex.Message); }

        }

        private string generarConsulta()
        {
            return txtConsulta.Text;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }
        private void btnElegirUsr_Click(object sender, EventArgs e)
        {
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                txtConsulta.Visible = true;
            }
            else
            {
                txtConsulta.Visible = false;
            }
        }
        string fileContent = string.Empty;
        string filePath = string.Empty;
        private void Button1_Click(object sender, EventArgs e)
        {
            string remoteUri = "http://www.banxico.org.mx/tipcamb/tipCamIHAction.do";
            string fileName = "tipoCambio.xls", myStringWebResource = null;
            if (DialogResult.Yes == MessageBox.Show("Descargar", "Si/No", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {

                if (File.Exists(fileName))
                {
                    try
                    {
                        File.Delete(fileName);

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("No se pudo abrir el archivo verifica que no este abierto.");return;
                    }
                }
                // Create a new WebClient instance.
                WebClient myWebClient = new WebClient();
                // Concatenate the domain with the Web resource filename.
                myStringWebResource = remoteUri;
                //Console.WriteLine("Downloading File \"{0}\" from \"{1}\" .......\n\n", fileName, myStringWebResource);
                // Download the Web resource and save it into the current filesystem folder.
                myWebClient.DownloadFile(myStringWebResource, fileName);

            }
            CargarEn(Application.StartupPath + "\\" + fileName, dgvResultado);

        }
        bool cargar = false;
        private string cargar1()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                    return filePath;
                }
            }
            return null;
        }
        public void CargarEn(string ruta, DataGridView dgv)
        {
            //Hoja desde donde obtendremos los datos
            string hoja = "Hoja1";
            if (cargar)
            {
                ruta = cargar1();
                if (ruta == null)
                {
                    return;
                }
            }


            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@ruta);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            for (int i = 0; i < colCount; i++)
            {
                dgv.Columns.Add("", "");
            }
            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= rowCount; i++)
            {
                label1.Text = "Renglon " + i + " De " + rowCount;
                dgv.Rows.Add();
                for (int j = 1; j <= colCount; j++)
                {
                    //new line
                    if (j == 1)
                        Console.Write("\r\n");
                    try
                    {
                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            dgv.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Value2.ToString();
                            int a = 0;
                            if (dgv.Rows[i - 1].Cells[0].Value != null && int.TryParse(dgv.Rows[i - 1].Cells[0].Value.ToString(), out a))
                            {
                                dgv.Rows[i - 1].Cells[0].Value = fecha.fechaRe(a);
                            }
                    }
                    catch { }
                }
            }

        //cleanup
        GC.Collect();
            GC.WaitForPendingFinalizers();
            //Cadena de conexión
            //string conexion = @"Provider=Microsoft.ACE.OLEDB.12.0;data source=" + @ruta + ";Extended Properties='Excel 12.0 Xml;HDR=Yes'";
            //string temp = ruta.Substring(ruta.Length - 3);
            //if (temp == "xls")
            //{
            //    conexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @ruta + ";Extended Properties = \"Excel 8.0;HDR=Yes;IMEX=1\"; ";
            //}
            //OleDbConnection con = new OleDbConnection(conexion);
            //try
            //{
            //    //throw new Exception("Test");
            //    //Conectarse al archivo de Excel
            //    con.Open();
            //    DataTable dbSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            //    if (dbSchema == null || dbSchema.Rows.Count < 1)
            //    {
            //        throw new Exception("Error: Could not determine the name of the first worksheet.");
            //    }
            //    string firstSheetName = dbSchema.Rows[0]["TABLE_NAME"].ToString();
            //    //Consulta contra la hoja de Excel
            //    OleDbCommand cmd = new OleDbCommand("Select * From [" + firstSheetName + "]", con);
            //    OleDbDataAdapter sda = new OleDbDataAdapter(cmd);
            //    DataTable data = new DataTable();

            //    //Cargar los datos
            //    sda.Fill(data);

            //    //Cargar la grilla
            //    dgv.Fill(data);

            //    foreach (DataGridViewRow row in dgv.Rows)
            //    {
            //        int a = 0;
            //        if (int.TryParse(row.Cells[0].Value.ToString(), out a))
            //        {
            //            row.Cells[0].Value = fecha.fechaRe(a);
            //        }
                    
            //    }
            //    //dgv.DataSource = data;
            //}
            //catch (Exception ex)
            //{
            //    //Error leyendo excel
            //    MessageBox.Show("Ocurrió un error en la lectura del archivo\n\n" + ex.Message);
            //    cargar = true;
            //    //CargarEn(ruta, dgv);
            //}
            //finally
            //{
            //    //Funcione o no, cerramos la cadena de conexión
            //    con.Close();
            //}
        }


        //De excell a DGV
        public void read(ref DataGridView dgv, string ruta)
        {

            dgv.Rows.Clear();
            //BuscarRuta(ref ruta, DatosPublicosGenerales.dirCortesSucursales);
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            object missing = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();

            xlWorkBook = xlApp.Workbooks.Open(ruta, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", true, true, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            range = xlWorkSheet.UsedRange;
            int i = 0;
            string str, str1;
            dgv.Rows.Clear();
            dgv.Columns.Clear();
            for (int x = 0; x < range.Columns.Count; x++)
            {
                dgv.Columns.Add(x + "", x + "");
            }
            for (int rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {
                dgv.Rows.Add();
                for (int cCnt = 1; cCnt < range.Columns.Count; cCnt++)
                {
                    string valor = "";
                    try
                    {
                        if (range.Cells[rCnt, cCnt] != null)
                        {
                            valor = (range.Cells[rCnt, cCnt] as Excel.Range).Cells.GetType().ToString();
                            valor = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2.ToString();

                        }
                        dgv.Rows[rCnt - 1].Cells[cCnt - 1].Value = valor;
                    }
                    catch (Exception)
                    {
                    }
                }

            }
            try
            {

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();


                GC.Collect();
                GC.WaitForPendingFinalizers();

                //release com objects to fully kill excel process from running in the background

                Marshal.ReleaseComObject(xlWorkSheet);

                //close and release
                //xlWorkBook.Close();
                Marshal.ReleaseComObject(xlWorkBook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception)
            {
            }
        }
        public void CargarEn1(string ruta, DataGridView dgv)
        {
            //Hoja desde donde obtendremos los datos
            string hoja = "Hoja1";


            //Cadena de conexión
            string conexion = @"Provider=Microsoft.ACE.OLEDB.12.0;data source=" + ruta + ";Extended Properties='Excel 12.0 Xml;HDR=Yes'";

            OleDbConnection con = new OleDbConnection(conexion);
            //Consulta contra la hoja de Excel
            OleDbCommand cmd = new OleDbCommand("Select * From [" + hoja + "$]", con);
            try
            {
                //Conectarse al archivo de Excel
                con.Open();

                OleDbDataAdapter sda = new OleDbDataAdapter(cmd);
                DataTable data = new DataTable();

                //Cargar los datos
                sda.Fill(data);

                //Cargar la grilla
                dgv.DataSource = data;
            }
            catch (Exception ex)
            {
                //Error leyendo excel
                MessageBox.Show("Ocurrió un error en la lectura del archivo\n\n" + ex.Message);
            }
            finally
            {
                //Funcione o no, cerramos la cadena de conexión
                con.Close();
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            con.ObtenerConexion();
        }
    }
}
