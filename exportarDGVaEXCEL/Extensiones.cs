using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace exportarDGVaEXCEL
{
    public static class Extensiones
    {
        /* String */
        public static int[] SplitToInt(this string ar, char car)
        {
            string[] arr = ar.Split(car);
            int[] arr2 = new int[arr.Count()];
            int i = 0;
            foreach (string valor in arr)
            {
                try
                {
                    arr2[i] = int.Parse(valor);
                }
                catch (Exception)
                {

                    throw new Exception("No se puede convertir el valor a entero. Valor{" + valor + "}");
                }
                i++;
            }
            return arr2;
        }

        /* Arreglos Arrays */
        public static object[] Add(this object[] arr, params object[] arr2)
        {
            object[] arrAux = new object[arr.Count() + arr2.Count()];
            int i = 0;
            foreach (object valor in arr)
            {
                try
                {
                    arrAux[i] = valor;
                }
                catch (Exception)
                {
                    throw new Exception("No se puede agregar el valor al arreglo. Valor{" + valor + "}");
                }
                i++;
            }
            foreach (object valor in arr2)
            {
                try
                {
                    arrAux[i] = valor;
                }
                catch (Exception)
                {
                    throw new Exception("No se puede agregar el valor al arreglo. Valor{" + valor + "}");
                }
                i++;
            }
            return arrAux;
        }

        /*public static object[] ToArray(this DataRow row)
        {
            object[] arr = new object[row.ItemArray.Count()];
            int i = 0;
            foreach (object item in row.ItemArray)
            {
                arr[i++] = item.ToString();
            }
            return arr;
        }*/


        /* DataGridView */
        public static void OrdenarDGV(this DataGridView dgv, params int[] arr)
        {
            DataGridView dgv2 = dgv;//FALTA
            foreach (DataGridViewRow row in dgv.Rows)
            {

            }
            for (int i = 0; i < arr.Length; i++)
            {

            }
        }

        public static void Fill(this DataGridView dgv, DataTable dt)
        {
            try
            {
                dgv.Rows.Clear();
                dgv.Columns.Clear();

            }
            catch (Exception)
            {

            }
            foreach (DataColumn column in dt.Columns)
            {
                dgv.Columns.Add(column.ColumnName, column.ColumnName);
            }
            foreach (DataRow dr in dt.Rows)
            {
                //string prueba = dr[5].ToString();
                dgv.Rows.Add(Array.ConvertAll(dr.ItemArray, ele => ele.ToString())); ;
            }
        }
        public static void FillOnlyColumns(this DataGridView dgv, DataTable dt, params int[] columnas)
        {
            try
            {
                dgv.Rows.Clear();
                dgv.Columns.Clear();

            }
            catch (Exception)
            {

            }
            foreach (int columnNum in columnas)
            {
                dgv.Columns.Add(dt.Columns[columnNum].ColumnName, dt.Columns[columnNum].ColumnName);
            }
            foreach (DataRow dr in dt.Rows)
            {
                string[] arr = new string[columnas.Count()];
                int i = 0;
                foreach (int columnNum in columnas)
                {
                    arr[i] = dr[columnNum].ToString();
                }
                dgv.Rows.Add(arr); ;
            }
        }


        public static void RowsAutoReSizeFill(this DataGridView dgv)
        {
            dgv.RowsAdded += new DataGridViewRowsAddedEventHandler(delegate (object sender, DataGridViewRowsAddedEventArgs e)
            {
                //
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    row.Height = (dgv.Height - dgv.ColumnHeadersHeight) / dgv.Rows.Count;
                }
            });
        }
        public static void RowsSizeFill(this DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.Height = (dgv.Height - dgv.ColumnHeadersHeight) / dgv.Rows.Count;
            }
        }
        /* DataTableCollection */
        public static DataTable First(this DataTableCollection dtc)
        {
            if (dtc.Count < 1)
            {
                return new DataTable();
            }
            else
            {
                if (dtc[0].Rows.Count < 1)
                {
                    return new DataTable();
                }
                return dtc[0];
            }
        }
        public static bool HasTables(this DataTableCollection dtc)
        {
            if (dtc.Count < 1)
            {
                return false;
            }
            else
            {
                //if (dtc[0].Rows.Count < 1)
                //{
                //    return false;
                //}
                return true;
            }
        }

        /* CheckBox */
        public static string ToStringS(this CheckBox rdb)
        {
            if (rdb.Checked)
            {
                return "S";
            }
            else
            {
                return "N"; ;
            }
        }
        public static void CheckedText(this CheckBox chk, Form th)
        {
            try
            {
                string nombre = chk.Name.Substring(3);
                Control ds = th.Controls.Find("txt" + nombre, true)?[0];

                if (chk.Checked)
                {
                    ((TextBox)ds).PasswordChar = '\0';
                    chk.CheckState = CheckState.Indeterminate;
                }
                else
                {
                    ((TextBox)ds).PasswordChar = '*';
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        /* DataTable */

        public static Dictionary<string, int> ToDictionary(this DataTable dt, string key, string value)
        {
            Dictionary<string, int> dict = new Dictionary<string, int>();
            foreach (DataRow row in dt.Rows)
            {
                try { dict.Add((string)row[key], (int)row[value]); } catch (Exception) { }
            }
            return dict;
        }

        /* Control */

        public static void OnlyNumbers(this Control ctrl)
        {
            Func<double, double> square = (double x) => { return x * x; };
            ctrl.KeyPress += new KeyPressEventHandler(delegate (object sender, KeyPressEventArgs e)
            {
                try
                {
                    if (Char.IsDigit(e.KeyChar))
                    {
                        e.Handled = false;
                    }
                    else if (Char.IsControl(e.KeyChar))
                    {
                        e.Handled = false;
                    }
                    else
                    {
                        e.Handled = true;
                    }
                }
                catch { }
            });

        }

        public static void OnlyLetters(this Control ctrl)
        {
            Func<double, double> square = (double x) => { return x * x; };
            ctrl.KeyPress += new KeyPressEventHandler(delegate (object sender, KeyPressEventArgs e)
            {
                if (Char.IsLetter(e.KeyChar))
                {
                    e.Handled = false;
                }
                else if (e.KeyChar == '-')
                {
                    e.Handled = false;
                }
                else if (e.KeyChar == Convert.ToChar("'"))
                {
                    e.Handled = false;
                }
                else if (Char.IsControl(e.KeyChar))
                {
                    e.Handled = false;
                }
                else if (Char.IsSeparator(e.KeyChar))
                {
                    e.Handled = false;
                }
                else
                {
                    e.Handled = true;
                }
            });
        }

        public static void OnlyNumbersNLetters(this Control ctrl)
        {
            Func<double, double> square = (double x) =>
            {
                return x * x;
            };
            ctrl.KeyPress += new KeyPressEventHandler(delegate (object sender, KeyPressEventArgs e)
            {
                if (Char.IsLetter(e.KeyChar))
                {
                    e.Handled = false;
                }
                else if (e.KeyChar == '-')
                {
                    e.Handled = false;
                }
                else if (e.KeyChar == Convert.ToChar("'"))
                {
                    e.Handled = false;
                }
                else if (Char.IsDigit(e.KeyChar))
                {
                    e.Handled = false;
                }
                else if (Char.IsControl(e.KeyChar))
                {
                    e.Handled = false;
                }
                else if (Char.IsSeparator(e.KeyChar))
                {
                    e.Handled = false;
                }
                else
                {
                    e.Handled = true;
                }
            });
        }
        public static void OnEnterDo(this Control txt, object btn)
        {
            txt.KeyDown += new KeyEventHandler(delegate (object sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter)
                {
                    ((Button)btn).PerformClick();
                }
            });
        }

        /* Controls */
        public static Control FindPanel(this Control.ControlCollection Controles, string nombre)
        {
            foreach (Control control in Controles)
            {
                if (control.Name == nombre)
                {
                    return control;
                }
            }
            return null;
        }
        /* TextBox */
        public static string text(this TextBox txt)
        {
            try
            {
                return txt.Text;
            }
            catch (Exception)
            {
            }
            return "";
        }

        /* DataRow */

        public static int ToInt(this object obj)
        {
            int _return = -1;
            if (obj is string)
            {
                if (obj.ToString() == "")
                {
                    obj = "0";
                }
                return int.Parse(obj.ToString());
            }
            else if (obj is bool) { return (bool)obj == true ? 1 : 0; }
            else
            {
                try
                {
                    return int.Parse(obj.ToString());
                }
                catch (Exception)
                {
                    MessageBox.Show("No se pudo convertir el valor " + obj + " a entero");
                }
            }
            return _return;
        }
        public static double ToDouble(this object obj)
        {
            double _return = -1;
            if (obj is string)
            {
                return double.Parse(obj.ToString());
            }
            else if (obj is bool) { return (bool)obj == true ? 1 : 0; }
            else
            {
                try
                {
                    return int.Parse(obj.ToString());
                }
                catch (Exception)
                {
                    MessageBox.Show("No se pudo convertir el valor " + obj + " a entero");
                }
            }
            return _return;
        }

        /* Form */
        public static string Preguntar(this Form fr, params string[] args)
        {
            //Argumento 0 es el mensaje
            //Argumento 1 es titulo del formulario
            //Argumento 2 si es password los caracteres son '*'
            string pregunta_descripcion = args.First();
            string titulo = args.Count() >= 2 ? args[1] : "Pregunta";

            Form pregunta = new Form(); pregunta.Size = new Size(400, 200); pregunta.StartPosition = FormStartPosition.CenterScreen; pregunta.Text = titulo;
            Label lbl = new Label(); lbl.Size = new Size(300, 20); lbl.Location = new Point(44, 28); lbl.Text = pregunta_descripcion;
            pregunta.Controls.Add(lbl);
            Button btn = new Button(); btn.Size = new Size(100, 30); btn.Location = new Point(230, 110); btn.Text = "OK";
            btn.Click += new EventHandler(delegate (object send, EventArgs e1)
            {
                pregunta.DialogResult = DialogResult.OK;
            });
            TextBox xtt = new TextBox(); xtt.Size = new Size(300, 20); xtt.Location = new Point(44, 50);
            xtt.OnEnterDo(btn);
            pregunta.Controls.Add(xtt);
            if (args.Count() >= 3 && args[2] == "password") { xtt.PasswordChar = '*'; }
            pregunta.Controls.Add(btn);
            pregunta.BringToFront();
            pregunta.ShowDialog();
            return xtt.text();
        }
        /* OpenFileDialog */

        public static object OpenFileD(params object[] args)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"C:\",
                Title = "Browse Image Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "jpg",
                Filter = "jpg files (*.jpg)|*.jpg|png files (*.png)|*.png|gif files (*.gif)|*.gif",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK && openFileDialog1.FileName != "")
            {
                return openFileDialog1.FileName;
            }
            return "";
        }
    }

    public class fecha
    {
        public static DateTime fechaRe(double dia)
        {
            DateTime fec = new DateTime(1899, 12, 30, 0, 0, 0);
            fec = fec.AddDays(dia);

            fec.ToShortDateString();

            return fec;
        }
    }
}