using Dapper;
using MailKit.Net.Smtp;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Timers;

namespace AlertaResina
{
    class Program
    {
        static  void Main(string[] args)
        {

            try
            {
                //var emailDto = new EmailDto();
                //emailDto.Asunto = "TEST";
                //emailDto.Mensaje = "mENSAJE DE PRUEBA";
                //List<string> list = new List<string>();
                //list.Add("edgar.gomez@te.com");
                //emailDto.Destinatarios = list.ToArray();
                //Enviar(emailDto);
                //var periodTimeSpan = TimeSpan.FromMinutes(5);


              //  Timer timer = new Timer();
                timer.Elapsed += new ElapsedEventHandler(RunConnection);
                timer.Interval = 30000; // 5 minutes in milliseconds
                timer.Enabled = true;

                Console.ReadKey();

            }
            catch (Exception ex)
            {
                var result = ex;
            }
            

           


        }
        //public static void PullFrame(object source, ElapsedEventArgs evArgs)
        //{
        //    // Do something every 5 mins.
        //}

        public static void RunConnection(object source, ElapsedEventArgs evArgs)
        {
            string connection = "Data Source= MXE41DB100;Initial Catalog=HYDRA_DATA;User ID=MXHydraUsrE41;Password=Password123";
            using (var db = new SqlConnection(connection))
            {
                var sql = "SELECT  [Cell_Name],[Factory_Order],[Part_Number],[Lot_Size],[Parts_Produced] FROM[HYDRA_DATA].[dbo].[Hydra_Data] where Value_Stream = 'DEU_MOLDING'  and Parts_Produced > Lot_Size";
                var lst = db.Query<HydraData>(sql);
                foreach (var oElement in lst)
                {
                    var existQuery = $"SELECT CASE WHEN EXISTS (SELECT* FROM[HYDRA_DATA].[dbo].[AlertasMoldeo] Where Factory_Order = '{oElement.Factory_Order}')THEN CAST(1 AS BIT) ELSE CAST(0 AS BIT) END";
                    var check = db.Query<bool>(existQuery);
                    if (!check.First())
                    {
                        var sqlInsert = "INSERT INTO [dbo].[AlertasMoldeo] ([Factory_Order],[Lot_Size],[Cell_Name],[Part_Number],[Parts_Produced]) VALUES(@Factory_Order, @Lot_Size, @Cell_Name, @Part_Number,@Parts_Produced)";
                        var result = db.Execute(sqlInsert, new { oElement.Factory_Order, oElement.Lot_Size, oElement.Cell_Name, oElement.Part_Number, oElement.Parts_Produced });

                        var emailDto = new EmailDto();
                        
                        //emailDto.Mensaje = @$"Alerta! Orden de moldeo ha exedido la piezas producidad de la orden
                        //                     PO:{oElement.Factory_Order}
                        //                     WC:{oElement.Cell_Name}   
                        //                     Part Number:{oElement.Part_Number}
                        //                     Lot size: {oElement.Lot_Size}
                        //                     Partes Producidas: {oElement.Parts_Produced}
                       // ";
                       emailDto.Mensaje = @$"<h2><span style='color: #ff0000;'>Alerta de Orden Excedida!</span></h2>
                        <p style='font-size: 1.5em;'>⚠La orden&nbsp;<strong style='background-color: #317399; padding: 0 5px; color: #fff;'>{oElement.Factory_Order}</strong> &nbsp; se ha excedido en la cantidad de piezas en maquina: <strong>{oElement.Cell_Name}</strong>⚠.</p>
               
                                                         <table class='editorDemoTable'>
                                          <tbody>
                                          <tr>
                                          <td>Orden: </td>
                                          <td>{oElement.Factory_Order.Substring(0, oElement.Factory_Order.Length-4)}</td>
                                          </tr>
                                          <tr>
                                          <td>WC: </td>
                                          <td>{oElement.Cell_Name}</td>

                                </tr>

                                <tr>

                                <td> PN: </td>

                                <td>{ oElement.Part_Number}</td>

                                </tr>

                                <tr>

                                <td>Lot Size: </td>

                                <td>{oElement.Lot_Size}<td>

                                </tr>

                                <tr>

                                <td> Producidas: </td>

                                <td>{oElement.Parts_Produced}</td>

                                </tr>

                                </tbody>

                                </table>

                                <p> &nbsp;</p>";
                        emailDto.Asunto = $"Orden Execedida: {oElement.Factory_Order} En WC:{oElement.Cell_Name}";
                        List<string> list = new List<string>();
                        list.Add("alertaordenesmoldeo@te.com");
                        emailDto.Destinatarios = list.ToArray();
                      Enviar(emailDto);

                    }
                    //Console.WriteLine(oElement.Cell_Name);
                }
            }
        }
        public static void Enviar(EmailDto emailDto)
        {

            var mailMessage = new MimeMessage();
            mailMessage.From.Add(new MailboxAddress("Servicio de envio de correos", "alertservice@te.com"));
            foreach (var destinatario in emailDto.Destinatarios)
            {
                mailMessage.To.Add(new MailboxAddress(destinatario, destinatario));
            }

            mailMessage.Subject = emailDto.Asunto;
            mailMessage.Body = new TextPart("html")
            {
                Text = "<b>" + emailDto.Mensaje + "</b>"
            };
            using (var smtpClient = new SmtpClient())
            {
                smtpClient.Connect("terelay.tycoelectronics.net", 25, MailKit.Security.SecureSocketOptions.None);
                smtpClient.Send(mailMessage);
                smtpClient.Disconnect(true);
            }


        }

    }
}
