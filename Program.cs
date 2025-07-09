using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Mail;
using Newtonsoft.Json.Linq;
using WOLF_START_MigrateDAR;
using Microsoft.Exchange.WebServices.Data;
using System.Text;
using System.IO;
using Attachment = System.Net.Mail.Attachment;

namespace SendEmail
{
    class Program
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(Program));

        private static string dbConnectionString
        {
            get
            {
                var ServarName = ConfigurationManager.AppSettings["ServarName"];
                var Database = ConfigurationManager.AppSettings["Database"];
                var Username_database = ConfigurationManager.AppSettings["Username_database"];
                var Password_database = ConfigurationManager.AppSettings["Password_database"];
                var dbConnectionString = $"data source={ServarName};initial catalog={Database};persist security info=True;user id={Username_database};password={Password_database};Connection Timeout=200";

                if (!string.IsNullOrEmpty(dbConnectionString))
                {
                    return dbConnectionString;
                }
                return "";
            }
        }
        private static string EsmtpServer
        {
            get
            {
                var smtpServer = ConfigurationManager.AppSettings["smtpServer"];
                if (!string.IsNullOrEmpty(smtpServer))
                {
                    return (smtpServer);
                }
                return string.Empty;
            }
        }
        private static int EsmtpPort
        {
            get
            {
                var smtpPort = ConfigurationManager.AppSettings["smtpPort"];
                if (!string.IsNullOrEmpty(smtpPort))
                {
                    return Convert.ToInt32(smtpPort);
                }
                return 0;
            }
        }
        private static string EsmtpUsername
        {
            get
            {
                var smtpUsername = ConfigurationManager.AppSettings["smtpUsername"];
                if (!string.IsNullOrEmpty(smtpUsername))
                {
                    return (smtpUsername);
                }
                return string.Empty;
            }
        }
        private static string EsmtpPassword
        {
            get
            {
                var smtpPassword = ConfigurationManager.AppSettings["smtpPassword"];
                if (!string.IsNullOrEmpty(smtpPassword))
                {
                    return (smtpPassword);
                }
                return string.Empty;
            }
        }
        private static string EfromEmail
        {
            get
            {
                var fromEmail = ConfigurationManager.AppSettings["fromEmail"];
                if (!string.IsNullOrEmpty(fromEmail))
                {
                    return (fromEmail);
                }
                return string.Empty;
            }
        }
        private static string EtoEmail
        {
            get
            {
                var toEmail = ConfigurationManager.AppSettings["toEmail"];
                if (!string.IsNullOrEmpty(toEmail))
                {
                    return (toEmail);
                }
                return string.Empty;
            }
        }
        private static int iIntervalTime
        {
            //ตั้งค่าเวลา
            get
            {
                var IntervalTime = ConfigurationManager.AppSettings["IntervalTimeMinute"];
                if (!string.IsNullOrEmpty(IntervalTime))
                {
                    return Convert.ToInt32(IntervalTime);
                }
                return -10;
            }
        }
        static void Main(string[] args)
        {
            try
            {
                log4net.Config.XmlConfigurator.Configure();
                log.Info("====== Start Process JobSendEmail ====== : " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                log.Info(string.Format("Run batch as :{0}", System.Security.Principal.WindowsIdentity.GetCurrent().Name));

                DbwolfDataContext db = new DbwolfDataContext(dbConnectionString);
                if (db.Connection.State == ConnectionState.Open)
                {
                    db.Connection.Close();
                    db.Connection.Open();
                }
                db.Connection.Open();
                db.CommandTimeout = 0;

                GetData(db);
            }
            catch (Exception ex)
            {
                Console.WriteLine(":ERROR");
                Console.WriteLine("exit 1");

                log.Error(":ERROR");
                log.Error("message: " + ex.Message);
                log.Error("Exit ERROR");
            }
            finally
            {
                log.Info("====== End Process Process JobSendEmail ====== : " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

            }
        }
        public static void GetData(DbwolfDataContext db)
        {
            var TemplateId = ConfigurationManager.AppSettings["TemplateId"];
            List<TRNMemo> lstmemo = db.TRNMemos.Where(x => x.TemplateId == Int32.Parse(TemplateId) && x.StatusName == "Wait for Approve" && x.ModifiedDate >= DateTime.Now.AddMinutes(iIntervalTime)).ToList();

            if (lstmemo.Count > 0)
            {
                foreach (var itemmemo in lstmemo) 
                {
                    var empseq1 = db.TRNLineApproves.Where(x => x.MemoId == itemmemo.MemoId && x.Seq == 1).Select(x => x.EmployeeId);
                    if (empseq1.Any() && empseq1.Contains(itemmemo.PersonWaitingId))
                    {
                        Console.WriteLine("Start Management Review : " + itemmemo.MemoId);
                        log.Info("Start Management Review : " + itemmemo.MemoId);
                        string Subject = "";
                        string TheTime = "";
                        string Date = "";
                        string TimeStart = "";
                        string TimeEnd = "";
                        string MeetingPlace = "";
                        string Annual = "";
                        string Standard = "";
                        List<object> listObjectArray = new List<object>();
                        List<object> Attendees = new List<object>();
                        List<object> Associate = new List<object>();
                        List<ViewEmployee> lstemp = new List<ViewEmployee>();
                        JObject jsonAdvanceForm = JsonUtils.createJsonObject(itemmemo.MAdvancveForm);
                        JArray itemsArray = (JArray)jsonAdvanceForm["items"];
                        foreach (JObject jItems in itemsArray)
                        {

                            JArray jLayoutArray = (JArray)jItems["layout"];
                            if (jLayoutArray.Count >= 1)
                            {
                                JObject jTemplateL = (JObject)jLayoutArray[0]["template"];
                                JObject jData = (JObject)jLayoutArray[0]["data"];
                                if ((String)jTemplateL["label"] == "เรื่อง")
                                {
                                    Subject = jData["value"].ToString();
                                }
                                if ((String)jTemplateL["label"] == "ครั้งที่")
                                {
                                    TheTime = jData["value"].ToString();
                                }
                                if ((String)jTemplateL["label"] == "วันที่")
                                {
                                    Date = jData["value"].ToString();
                                }
                                if ((String)jTemplateL["label"] == "เวลาเริ่ม")
                                {
                                    TimeStart = jData["value"].ToString();
                                }
                                if ((String)jTemplateL["label"] == "สถานที่ประชุม")
                                {
                                    MeetingPlace = jData["value"].ToString();
                                }
                                if ((String)jTemplateL["label"] == "รายนามผู้เข้าร่วมประชุม")
                                {

                                    foreach (JArray row in jData["row"])
                                    {
                                        List<object> rowObject = new List<object>();
                                        foreach (JObject item in row)
                                        {
                                            rowObject.Add(item["value"].ToString());
                                        }
                                        listObjectArray.Add(rowObject);
                                        Attendees.Add(rowObject);
                                    }
                                }
                                if ((String)jTemplateL["label"] == "รายนามผู้เข้าร่วมประชุมสมทบ")
                                {

                                    foreach (JArray row in jData["row"])
                                    {
                                        List<object> rowObject = new List<object>();
                                        foreach (JObject item in row)
                                        {
                                            rowObject.Add(item["value"].ToString());
                                        }
                                        listObjectArray.Add(rowObject);
                                        Associate.Add(rowObject);
                                    }
                                }
                                if (jLayoutArray.Count > 1)
                                {
                                    JObject jTemplateR = (JObject)jLayoutArray[1]["template"];
                                    JObject jData2 = (JObject)jLayoutArray[1]["data"];
                                    if ((String)jTemplateR["label"] == "ประจำปี")
                                    {
                                        Annual = jData2["value"].ToString();
                                    }
                                    if ((String)jTemplateR["label"] == "เวลาจบ")
                                    {
                                        TimeEnd = jData2["value"].ToString();
                                    }
                                    if ((String)jTemplateR["label"] == "Standard")
                                    {
                                        Standard = jData2["value"].ToString();
                                    }
                                }
                            }
                        }
                        if (listObjectArray.Count() > 0)
                        {
                            foreach (var Eitem in listObjectArray)
                            {
                                dynamic item = Eitem;
                                string Empolyeecode = item[0];
                                string PositionMeeting = item[3];
                                string AttendMeeting = item[4];
                                if (!Empolyeecode.ToLower().Contains("select"))
                                {
                                    ViewEmployee emp = db.ViewEmployees.Where(e => e.EmployeeCode.Contains(Empolyeecode)).FirstOrDefault();
                                    lstemp.Add(emp);
                                }

                            }
                            if (lstemp.Count > 0)
                            {
                                fSendCalendars(lstemp, Subject, Date, TimeStart, TimeEnd, itemmemo, MeetingPlace, db, Annual, TheTime, Standard, Attendees, Associate);
                            }
                        }
                        log.Info("----------------------------------------------------------------------------");
                    }
                }
            }
            else
            {
                Console.WriteLine("Management Review : " + lstmemo.Count);
                log.Info("Management Review : " + lstmemo.Count);
            }
        }
        public static void fSendCalendars(List<ViewEmployee> emp, string Subject, string Date, string TimeStart, string TimeEnd, TRNMemo itemmemo, string MeetingPlace, DbwolfDataContext db, string Annual, string TheTime, string Standard, List<object> Attendees, List<object> Associate)
        {
            string AllAttendees = "";
            if (Attendees.Count() > 0)
            {
                StringBuilder sb = new StringBuilder();
                foreach (var Eitem in Attendees)
                {
                    dynamic item = Eitem;
                    string name = item[1];

                    if (sb.Length > 0)
                    {
                        sb.Append(", ");
                    }
                    sb.Append(name);
                }
                AllAttendees = sb.ToString();
            }
            string AllAssociate = "";
            if (Associate.Count() > 0)
            {
                StringBuilder sb = new StringBuilder();
                foreach (var Eitem in Associate)
                {
                    dynamic item = Eitem;
                    string name = item[1];

                    if (sb.Length > 0)
                    {
                        sb.Append(", ");
                    }
                    sb.Append(name);
                }
                AllAssociate = sb.ToString();
            }

            var TemplateId = ConfigurationManager.AppSettings["TemplateId"];
            MSTEmailTemplate emailtem = db.MSTEmailTemplates.Where(e => e.TemplateId == Int32.Parse(TemplateId)).FirstOrDefault();

            DateTime date = DateTime.ParseExact(Date, "dd MMM yyyy", System.Globalization.CultureInfo.InvariantCulture);
            DateTime startTime = DateTime.ParseExact(TimeStart, "h:mm tt", System.Globalization.CultureInfo.InvariantCulture);
            DateTime endTime = DateTime.ParseExact(TimeEnd, "h:mm tt", System.Globalization.CultureInfo.InvariantCulture);

            DateTime startDateTime = new DateTime(date.Year, date.Month, date.Day, startTime.Hour, startTime.Minute, 0);
            DateTime endDateTime = new DateTime(date.Year, date.Month, date.Day, endTime.Hour, endTime.Minute, 0);
            MailMessage msg = new MailMessage();
            string body = emailtem.EmailBody;

            body = body.Replace("[DearName]", "Approver")
            .Replace("[Subject]", Subject)
            .Replace("[Thetime]", TheTime)
            .Replace("[Year]", Annual)
            .Replace("[Date]", date.ToString("dd/MM/yyyy"))
            .Replace("[startDateTime-endDateTime]", $"{startDateTime}-{endDateTime}")
            .Replace("[MeetingPlace]", MeetingPlace)
            .Replace("[Standard]", Standard)
            .Replace("[Attendees]", AllAttendees)
            .Replace("[Associate]", AllAssociate);

            if (!string.IsNullOrEmpty(EtoEmail))
            {
                msg = new MailMessage(EfromEmail, EtoEmail)
                {
                    Subject = Subject,
                    Body = body,
                    IsBodyHtml = true
                };


                StringBuilder str = new StringBuilder();
                str.AppendLine("BEGIN:VCALENDAR");
                str.AppendLine("PRODID:-//Schedule a Meeting");
                str.AppendLine("VERSION:2.0");
                str.AppendLine("METHOD:REQUEST");
                str.AppendLine("BEGIN:VEVENT");
                str.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", startDateTime.ToUniversalTime()));
                str.AppendLine(string.Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", DateTime.UtcNow));
                str.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", endDateTime.ToUniversalTime()));
                str.AppendLine($"LOCATION: {MeetingPlace}");
                str.AppendLine(string.Format("UID:{0}", Guid.NewGuid()));
                str.AppendLine(string.Format("SUMMARY:{0}", msg.Subject));
                str.AppendLine(string.Format("ORGANIZER:MAILTO:{0}", msg.From.Address));
                str.AppendLine(string.Format("ATTENDEE;CN=\"{0}\";RSVP=TRUE:mailto:{1}", "กำปั่น", "thearaphat@techconsbiz.com"));
                str.AppendLine("BEGIN:VALARM");
                str.AppendLine("TRIGGER:-PT15M");
                str.AppendLine("ACTION:DISPLAY");
                str.AppendLine("DESCRIPTION:Reminder");
                str.AppendLine("END:VALARM");
                str.AppendLine("END:VEVENT");
                str.AppendLine("END:VCALENDAR");

                byte[] byteArray = Encoding.UTF8.GetBytes(str.ToString());
                MemoryStream stream = new MemoryStream(byteArray);
                Attachment attach = new Attachment(stream, "meeting.ics", "text/calendar; charset=UTF-8");

                var FilePath_ICS = ConfigurationManager.AppSettings["FilePath_ICS"];
                string filePath = $@"{FilePath_ICS}";
                using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    stream.WriteTo(fileStream);
                }
                msg.Attachments.Add(attach);
                msg.BodyEncoding = Encoding.UTF8;
                msg.SubjectEncoding = Encoding.UTF8;

                SmtpClient smtpClient = new SmtpClient(EsmtpServer, EsmtpPort);
                smtpClient.EnableSsl = false;
                if (!string.IsNullOrEmpty(EsmtpUsername) && !string.IsNullOrEmpty(EsmtpPassword))
                {
                    smtpClient.Credentials = new NetworkCredential(EsmtpUsername, EsmtpPassword);
                    smtpClient.EnableSsl = true;
                }
                try
                {
                    smtpClient.Send(msg);
                    Console.WriteLine("SendCalendars Successful. DocumentNo: " + itemmemo.DocumentNo + " && Email to : " + EtoEmail);
                    log.Info("SendCalendars Successful. DocumentNo: " + itemmemo.DocumentNo + " && Email to : " + EtoEmail);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("SendCalendars Error: " + ex.Message + " || DocumentNo: " + itemmemo.DocumentNo);
                    log.Info("SendCalendars Error: " + ex.Message + " || DocumentNo: " + itemmemo.DocumentNo);
                }
            }
            else
            {
                foreach (var employee in emp)
                {
                    msg = new MailMessage(EfromEmail, employee.Email)
                    {
                        Subject = Subject,
                        Body = body,
                        IsBodyHtml = true
                    };
                    StringBuilder str = new StringBuilder();
                    str.AppendLine("BEGIN:VCALENDAR");
                    str.AppendLine("PRODID:-//Schedule a Meeting");
                    str.AppendLine("VERSION:2.0");
                    str.AppendLine("METHOD:REQUEST");
                    str.AppendLine("BEGIN:VEVENT");
                    str.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", startDateTime.ToUniversalTime()));
                    str.AppendLine(string.Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", DateTime.UtcNow));
                    str.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", endDateTime.ToUniversalTime()));
                    str.AppendLine($"LOCATION: {MeetingPlace}");
                    str.AppendLine(string.Format("UID:{0}", Guid.NewGuid()));
                    str.AppendLine(string.Format("SUMMARY:{0}", msg.Subject));
                    str.AppendLine(string.Format("ORGANIZER:MAILTO:{0}", msg.From.Address));

                    foreach (var Aemployee in emp)
                    {
                        str.AppendLine(string.Format("ATTENDEE;CN=\"{0}\";RSVP=TRUE:mailto:{1}", Aemployee.NameTh, Aemployee.Email));
                    }

                    str.AppendLine("BEGIN:VALARM");
                    str.AppendLine("TRIGGER:-PT15M");
                    str.AppendLine("ACTION:DISPLAY");
                    str.AppendLine("DESCRIPTION:Reminder");
                    str.AppendLine("END:VALARM");
                    str.AppendLine("END:VEVENT");
                    str.AppendLine("END:VCALENDAR");

                    byte[] byteArray = Encoding.UTF8.GetBytes(str.ToString());
                    MemoryStream stream = new MemoryStream(byteArray);
                    Attachment attach = new Attachment(stream, "meeting.ics", "text/calendar; charset=UTF-8");

                    var FilePath_ICS = ConfigurationManager.AppSettings["FilePath_ICS"];
                    string filePath = $@"{FilePath_ICS}";
                    using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                    {
                        stream.WriteTo(fileStream);
                    }
                    msg.Attachments.Add(attach);
                    msg.BodyEncoding = Encoding.UTF8;
                    msg.SubjectEncoding = Encoding.UTF8;

                    SmtpClient smtpClient = new SmtpClient(EsmtpServer, EsmtpPort);
                    smtpClient.EnableSsl = false;
                    if (!string.IsNullOrEmpty(EsmtpUsername) && !string.IsNullOrEmpty(EsmtpPassword))
                    {
                        smtpClient.Credentials = new NetworkCredential(EsmtpUsername, EsmtpPassword);
                        smtpClient.EnableSsl = true;
                    }
                    try
                    {
                        smtpClient.Send(msg);
                        Console.WriteLine("SendCalendars Successful. DocumentNo: " + itemmemo.DocumentNo + " && Email to : " + employee.Email);
                        log.Info("SendCalendars Successful. DocumentNo: " + itemmemo.DocumentNo + " && Email to : " + employee.Email);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("SendCalendars Error: " + ex.Message + " || DocumentNo: " + itemmemo.DocumentNo);
                        log.Info("SendCalendars Error: " + ex.Message + " || DocumentNo: " + itemmemo.DocumentNo);
                    }
                }
            }
        }
    }
}
