using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace exportarDGVaEXCEL
{
    public class Conexion
    {
        /* MySQL */
        public MySqlCommand MySqlCmd = new MySqlCommand();

        static public MySqlConnection MySqlCnn { get; set; }

        public MySqlConnection ObtenerMySQLConnection() { if (MySqlCnn.ConnectionString == "") { ObtenerConexion(); }  return MySqlCnn; }
        public MySqlConnection ObtenerConexion()
        {
            MySqlCnn = new MySqlConnection("" +
                "Server=" + new Form().Preguntar("IP","Ingresa la dirección de IP")+ "; " + 
                "Port=3306;" + 
                "Database="+ new Form().Preguntar("Base de Datos", "Ingresar nombre de base de datos")+";" + 
                "Uid=" + new Form().Preguntar("usuario", "Ingresar usuario") + ";" + 
                "Password=" + new Form().Preguntar("Contraseña","Ingresa la contraseña") + 
                "; Convert Zero Datetime=True; Command Timeout=128800;");
            return MySqlCnn;
        }
        public MySqlConnection ObtenerConexionLocalHost()
        {
            //MySqlCnn = DatosCajeroSucursalOperacion.MySqlCnn;
            MySqlCnn = new MySqlConnection("Server=" + "localhost" + "; " + "Port=3306;" + "Uid=" + "root" + ";" + "Password=" + new Form().Preguntar("Contraseña","Ingresa la contraseña") + "; Convert Zero Datetime=True; Command Timeout=128800;");
            return MySqlCnn;
        }
        public Conexion()
        {
            ObtenerConexion();
        }
        public DataSet CargarSelect(string CON)
        {
            DataSet dt = new DataSet();
            try
            {
                string ConsultaProductos = CON; // Definimos nuestro Query
                MySqlCmd = new MySqlCommand(ConsultaProductos, MySqlCnn);
                MySqlDataAdapter da = new MySqlDataAdapter(MySqlCmd);
                da.Fill(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                MySqlCnn.Close();
            }
            return dt;
        }
        public static void GuardarImagen(Image imagen, string ID)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                imagen.Save(ms, ImageFormat.Gif);
                byte[] imgArr = ms.ToArray();
                using (MySqlConnection conn = new Conexion().ObtenerConexion())
                {
                    using (MySqlCommand cmd = new MySqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "INSERT INTO ;";
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.Parameters.AddWithValue("@Image", imgArr);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }
        public static Image CargarImagen(string ID)
        {
            try
            {

                using (MySqlConnection conn = new Conexion().ObtenerConexion())
                {
                    using (MySqlCommand cmd = new MySqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = "SELECT Image FROM  WHERE ID = @ID";
                        cmd.Parameters.AddWithValue("@ID", ID);
                        byte[] imgArr = (byte[])cmd.ExecuteScalar();
                        imgArr = (byte[])cmd.ExecuteScalar();
                        using (var stream = new MemoryStream(imgArr))
                        {
                            Image img = Image.FromStream(stream);
                            return img;
                        }
                    }
                }
            }
            catch (Exception)
            {
                //MessageBox.Show("No se pudo obtener la imagen");
                return null;
            }
        }
        /* XML */
        private XElement xmlContactos;
        public void LeerXML()
        {
            try
            {
                string path = @"C:\Path";
                string Datos = path + @"\File.xml";
                if (!File.Exists(Datos))
                {
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                }
                XmlDocument xDoc = new XmlDocument();
                //La ruta del documento XML permite rutas relativas 
                xDoc.LoadXml("<XML FILE IN FORMAT....>");
                XmlNodeList configuracion = xDoc.GetElementsByTagName("configuraciones");
                XmlNodeList lista = ((XmlElement)configuracion[0]).GetElementsByTagName("configuracion");
                foreach (XmlElement nodo in lista)
                {
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }
}
