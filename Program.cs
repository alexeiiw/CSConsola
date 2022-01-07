using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics; 

namespace Monitoreo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Define variable para el loop
            bool salir = false;

            do
            {
                // Se ejecutara cada 15 minutos para verificar si la interfaz esta funcionando
                Thread.Sleep(900000);
                //Thread.Sleep(300000);

                GenerarForecastSISCON(ref salir);

                EjecutarEnvioCorreosLecturas(ref salir); 

                RealizarMonitoreoSISCON(ref salir);

                RealizarMonitoreoSIFCO(ref salir);

                RealizarMonitoreoSISCONDespachos(ref salir);

                RealizarMonitoreoSISCONRetiros(ref salir);

                RealizarMonitoreoSISCONDespachosNoContratos(ref salir);

                RealizarMonitoreoSISCONInconsistencias(ref salir);
            }
            while (salir == false);
        }

        static private void RealizarMonitoreoSISCONInconsistencias(ref bool salir)
        {
            string connStringUtil;

            connStringUtil = "data source=128.1.200.167;initial catalog=Canella_SISCON;persist security info=True;user id=usrsap;password=C@nella20$";

            string query = @"select itemcode from SBO_CANELLA.dbo.OITM where U_Condicion = 'D' and itemcode in 
                            (select assetid_sap from COT_CONTRATOS_DET_DESPACHO where ESTATUS_ARTICULO = 1)
                            and ItemCode <> '1' order by ItemCode";

            DataTable dt = new DataTable();

            using (SqlConnection connUtil = new SqlConnection(connStringUtil))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(query, connUtil))
                {
                    connUtil.Open();

                    da.Fill(dt);

                    connUtil.Close();
                }
            }

            if (dt.Rows.Count > 0)
            {
                Console.WriteLine("Hay incosistencia de datos SISCON versus SAP " + "Alerta Inconsistencias de datos SISCON");
                EnviarCorreoInternos("aarrecis@canella.com.gt","Hay incosistencia de datos SISCON versus SAP ", "Alerta Inconsistencias de datos SISCON");
                salir = true;
            }
            else
            {
                Console.WriteLine("No hay incosistencia de datos SISCON versus SAP");
            }
        }

        static private void RealizarMonitoreoSIFCO(ref bool salir)
        {
            string connStringUtil;

            connStringUtil = "data source=128.1.200.169;initial catalog=canella_crp;persist security info=True;user id=usrsap;password=C@nella20$";

            string query = @"select * from SIFCO.TExternalCallOutTransactions 
                            where ExternalCallOutTransactions_Estado = 3";

            DataTable dt = new DataTable();

            using (SqlConnection connUtil = new SqlConnection(connStringUtil))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(query, connUtil))
                {
                    connUtil.Open();

                    da.Fill(dt);

                    connUtil.Close();
                }
            }

            if (dt.Rows.Count > 0)
            {
                Console.WriteLine("La interfaz se ha detenido por favor verificar " + "Alerta de Interfaz SAP - SIFCO");
                EnviarCorreoInternos("aarrecis@canella.com.gt","La interfaz se ha detenido por favor verificar ", "Alerta de Interfaz SAP - SIFCO");
                salir = true;
            }
            else
            {
                Console.WriteLine("La interfaz SIFCO esta funcionando correctamente");
            }
        }

        static private void RealizarMonitoreoSISCONRetiros(ref bool salir)
        {
            string connStringUtil;

            connStringUtil = "data source=128.1.200.167;initial catalog=Canella_SISCON;persist security info=True;user id=usrsap;password=C@nella20$";

            string query = @"select u_condicion, u_contracid, itemcode from SBO_CANELLA.dbo.OITM where itemcode in (
                            select ASSETID_SAP from Canella_SISCON.dbo.cot_contratos_det_despacho where estatus_articulo = 2 and ASSETID_SAP <> '0' and ASSETID_SAP <> '1' and ASSETID_SAP not in 
                            (
                            select ASSETID_SAP from Canella_SISCON.dbo.cot_contratos_det_despacho where estatus_articulo = 1 and ASSETID_SAP <> '0' and ASSETID_SAP <> '1'
                            )) and u_condicion = 'A' and SUBSTRING(itemcode,1,2) <> 'ME'";

            DataTable dt = new DataTable();

            using (SqlConnection connUtil = new SqlConnection(connStringUtil))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(query, connUtil))
                {
                    connUtil.Open();

                    da.Fill(dt);

                    connUtil.Close();
                }
            }

            if (dt.Rows.Count > 0)
            {
                Console.WriteLine("Existen retiros que no se actualizaron en SAP " + "Alerta de Retiros SISCON");
                EnviarCorreoInternos("aarrecis@canella.com.gt","Existen retiros que no se actualizaron en SAP ", "Alerta de Retiros SISCON");
                salir = true;
            }
            else
            {
                Console.WriteLine("Los retiros SISCON estan funcionando correctamente");
            }
        }

        static private void RealizarMonitoreoSISCONDespachos(ref bool salir)
        {
            string connStringUtil;

            connStringUtil = "data source=128.1.200.167;initial catalog=Canella_SISCON;persist security info=True;user id=usrsap;password=C@nella20$";

            string query = @"select * from COT_DESPACHO_ENC where ENTREGA_ESTADO = 1 and ESTADO_SBO = 0";

            DataTable dt = new DataTable();

            using (SqlConnection connUtil = new SqlConnection(connStringUtil))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(query, connUtil))
                {
                    connUtil.Open();

                    da.Fill(dt);

                    connUtil.Close();
                }
            }

            if (dt.Rows.Count > 0)
            {
                Console.WriteLine("Existen despachos que no generan ordenes de venta en SAP " + "Alerta de Despachos SISCON");
                EnviarCorreoInternos("aarrecis@canella.com.gt","Existen despachos que no generan ordenes de venta en SAP ", "Alerta de Despachos SISCON");
                salir = true;
            }
            else
            {
                Console.WriteLine("Los despachos SISCON estan funcionando correctamente");
            }
        }

        static private void RealizarMonitoreoSISCONDespachosNoContratos(ref bool salir)
        {
            string connStringUtil;

            connStringUtil = "data source=128.1.200.167;initial catalog=Canella_SISCON;persist security info=True;user id=usrsap;password=C@nella20$";

            string query = @"select * from (
                            select DES.DESPACHO_NO Despacho, DES.ASSETID_SAP Codigo_Despacho, DES.LINEA_CONTRATO, DES.COD_ARTICULO, DES.NUMERO_SERIE, DES.TIPO_LINEA,  
                            DES.ARTICULO_PADRE, DES.LINEA_PADRE, DES.CONTRATO_NO ContratoDespacho, 
                            COT.ESTATUS_ARTICULO, COT.ASSETID_SAP Codigo_Contrato, COT.CONTRATO_NO ContratoContrato
                            from COT_DESPACHO_DET DES left join COT_CONTRATOS_DET_DESPACHO COT on DES.ASSETID_SAP = COT.ASSETID_SAP 
                            and DES.CONTRATO_NO = COT.CONTRATO_NO
                            where DES.DESPACHO_NO in (
                            select a.DESPACHO_NO
                            from COT_DESPACHO_ENC a
                            where a.FECHA_CREADO > '2021-01-01'
                            and a.ENTREGA_ESTADO = 2
                            )) as datos 
                            where Codigo_Contrato is null";

            DataTable dt = new DataTable();

            using (SqlConnection connUtil = new SqlConnection(connStringUtil))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(query, connUtil))
                {
                    connUtil.Open();

                    da.Fill(dt);

                    connUtil.Close();
                }
            }

            if (dt.Rows.Count > 0)
            {
                Console.WriteLine("Existen despachos que no crearon articulos en los Contratos de SISCON y tampoco estan actualizados en AF SAP " + "Alerta de Despachos SISCON");
                EnviarCorreoInternos("aarrecis@canella.com.gt","Existen despachos que no crearon articulos en los Contratos de SISCON y tampoco estan actualizados en AF SAP ", "Alerta de Despachos SISCON");
                salir = true;
            }
            else
            {
                Console.WriteLine("Los despachos SISCON y su relación en Contratos estan funcionando correctamente");
            }
        }

        static private void GenerarForecastSISCON(ref bool salir)
        {
            string connStringUtil;

            connStringUtil = "data source=128.1.200.167;initial catalog=Canella_SISCON;persist security info=True;user id=usrsap;password=C@nella20$";

            string query = "select fecha_generado from REP_CONTRATOS_FORECAST";

            DataTable dt = new DataTable();

            using (SqlConnection connUtil = new SqlConnection(connStringUtil))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(query, connUtil))
                {
                    connUtil.Open();

                    da.Fill(dt);

                    connUtil.Close();
                }
            }
            
            // Obtiene la última fecha de generación del insumo y lo compara con la fecha actual
            int intCompara = DateTime.Compare(DateTime.Today, Convert.ToDateTime(dt.Rows[0]["fecha_generado"]));

            // Si es más tarde
            if (intCompara > 0)
            {
                try
                {
                    using (SqlConnection connUtil = new SqlConnection(connStringUtil)) 
                    {
                        using (SqlCommand cmd = new SqlCommand("Generar_REP_CONTRATOS_FORECAST", connUtil))    
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.CommandTimeout = 600;

                            connUtil.Open();

                            cmd.ExecuteNonQuery();

                            connUtil.Close();
                        }
                    }

                    Console.WriteLine("Se ha generado de forma correcta el reporte para Forecast SISCON " + "Información de Generación - SISCON");

                    EnviarCorreoInternos("aarrecis@canella.com.gt","Se ha generado de forma correcta el reporte para Forecast SISCON ", "Información de Generación - SISCON");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.ToString());

                    EnviarCorreoInternos("aarrecis@canella.com.gt",e.ToString(), "Alerta de Generación Reporte Forecast - SISCON");

                    salir = true;
                }
            }
            else { Console.WriteLine("No es necesario generar el reporte de Forecast " + "Información de Generación - SISCON"); }
        }

        static private void RealizarMonitoreoSISCON(ref bool salir)
        {
            // Obtiene las fechas del archivo
            //DateTime fileCreatedDate = File.GetLastWriteTime(@"C:\Program Files (x86)\IGT\069 - Interfaz SAP B1 Activos Fijos y Sistema Contratos SISCON Canella\logs\Consola\app-log.txt"); // 32 bits
            DateTime fileCreatedDate = File.GetLastWriteTime(@"C:\Program Files\IGT\069 - Interfaz SAP B1 Activos Fijos y Sistema Contratos SISCON Canella\logs\Consola\app-log.txt"); // 64 bits
            DateTime fecAhora = DateTime.Now;

            Console.WriteLine("El archivo fue modificado a las: " + fileCreatedDate.ToString());
            Console.WriteLine("La fecha de hoy para comparar es: " + fecAhora.ToString());

            // Obtiene la diferencia de tiempo entre la modificación del archivo y la fecha de ahora
            var objTiempo = fecAhora - fileCreatedDate;

            Console.WriteLine("La diferencia es: " + objTiempo.Minutes.ToString() + " Minutos");

            // Realiza la validación, verifica que al menos el archivo se haya modificado en los ultimos 5 minutos
            if (objTiempo.Minutes > 5)
            {
                Console.WriteLine("La interfaz se ha detenido por favor verificar " + "Alerta de Interfaz SAP - Activos Fijos");
                EnviarCorreoInternos("aarrecis@canella.com.gt","La interfaz se ha detenido por favor verificar ", "Alerta de Interfaz SAP - Activos Fijos");
                salir = true;
            }
            else
            {
                Console.WriteLine("La interfaz SISCON esta funcionando correctamente");
            }
        }

        private static void EjecutarEnvioCorreosLecturas(ref bool salir)
        {
            string connStringLecturas;

            connStringLecturas = "data source=128.1.200.167;initial catalog=Lecturas;persist security info=True;user id=usrsap;password=C@nella20$";

            string query = "select Valor from LECT_Configuracion where Nombre = 'Fecha_Generado'";

            DataTable dt = new DataTable();

            using (SqlConnection connLecturas = new SqlConnection(connStringLecturas))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(query, connLecturas))
                {
                    connLecturas.Open();

                    da.Fill(dt);

                    connLecturas.Close();
                }
            }

            // Obtiene la última fecha de generación del insumo y lo compara con la fecha actual
            int intCompara = DateTime.Compare(DateTime.Today, Convert.ToDateTime(dt.Rows[0]["valor"]));

            // Si es más tarde
            if (intCompara > 0)
            {
                Process myProcess = new Process();

                try
                {
                    myProcess.StartInfo.UseShellExecute = false;
                    // You can start any process, HelloWorld is a do-nothing example.
                    myProcess.StartInfo.FileName = "e:\\LecturasRPT\\ServicioTecnicoReporte.exe";
                    myProcess.StartInfo.CreateNoWindow = true;
                    myProcess.Start();
                    // This code assumes the process you are starting will terminate itself.
                    // Given that is is started without a window so you cannot terminate it
                    // on the desktop, it must terminate itself or you can do it programmatically
                    // from this application using the Kill method.

                    Console.WriteLine("Se ha generado en un hilo diferente el envío de correos y archivos - LECTURAS");

                    ActualizarFechaCorreoLecturas();

                    EnviarCorreoInternos("aarrecis@canella.com.gt","Se ha iniciado el proceso para envío de  correos y archivos de LECTURAS ", "Información de Generación - LECTURAS");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);

                    EnviarCorreoInternos("aarrecis@canella.com.gt",e.ToString(), "Alerta de Generación Proceso Correos y Archivos - LECTURAS");

                    salir = true;
                }        
            }
            else { Console.WriteLine("No es necesario generar el proceso de envío de correos de Lecturas " + "Información de Generación - LECTURAS"); }
        }

        private static void ActualizarFechaCorreoLecturas()
        {
            string connStringLecturas;

            connStringLecturas = "data source=128.1.200.167;initial catalog=Lecturas;persist security info=True;user id=usrsap;password=C@nella20$";

            using (SqlConnection connLecturas = new SqlConnection(connStringLecturas))
            {
                connLecturas.Open();

                SqlCommand command = connLecturas.CreateCommand();

                command.Connection = connLecturas;

                string strUpdate = "update LECT_Configuracion set Valor = GETDATE() where Nombre = 'Fecha_Generado'";
                
                command.CommandText = strUpdate;

                command.ExecuteNonQuery();

                connLecturas.Close();
            }
        }

        private static void EnviarCorreoInternos(string ToMail, string BodyMail, string SubjectMail)
        {
            string ServidorCorreoCanella = "srv-ex2010";

            MailMessage mail = new MailMessage();
            mail.From = new MailAddress("alertassap@canella.com.gt", "Sistema Monitoreo - Alerta");
            mail.To.Add(ToMail);
            mail.Subject = SubjectMail;
            mail.Body = BodyMail;
            mail.IsBodyHtml = true;
            //SmtpClient smpt = new SmtpClient("128.1.200.141");
            SmtpClient smpt = new SmtpClient(ServidorCorreoCanella);
            //            smpt.Port = 25;
            //            smpt.UseDefaultCredentials = true;
            //            smpt.Timeout = 25;
            smpt.Send(mail);
        }

        private static void EmailAnterior(string htmlString, string subjString)
        {
            try
            {
                MailMessage message = new MailMessage();
                SmtpClient smtp = new SmtpClient();
                message.From = new MailAddress("alexeiiw@hotmail.com");
                message.To.Add(new MailAddress("aarrecis@canella.com.gt"));
                message.Subject = subjString;
                message.IsBodyHtml = true; //to make message body as html  
                message.Body = htmlString;
                smtp.Port = 587;
                smtp.Host = "smtp.live.com"; //for hotmail host  
                smtp.EnableSsl = true;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new NetworkCredential("alexeiiw@hotmail.com", "ag19231923");
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Send(message);
            }
            catch (Exception) { }
        }
    }
}
