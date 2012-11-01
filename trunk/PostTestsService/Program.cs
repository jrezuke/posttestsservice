using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NLog;
using System.Configuration;
using System.Net;
using System.Net.Mail;
using System.IO;
using System.Web;
using System.Web.Security;

namespace PostTestsService
{
    class Program
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        
        static void Main(string[] args)
        {           

            logger.Info("Starting PostTests Service");
            
            string path = System.AppDomain.CurrentDomain.BaseDirectory;
                       
            //get sites 
            var sites = GetSites();
            
            //iterate sites
            foreach (var si in sites)
            {
                var ptndcl = new List<PostTestNextDue>();
                var ptndc2l = new List<PostTestNextDue>();
                var ptndc3l = new List<PostTestNextDue>();
                var ptndc4l = new List<PostTestNextDue>();

                Console.WriteLine(si.Name);
                logger.Info("For Site:" + si.Name + " - " + si.SiteID);
                logger.Debug("For Site:" + si.Name + " - " + si.SiteID);
                //Get the next date due for people - this works on tests
                //that are current (IsCurrent=1)
                var ptndl = GetPostTestPeopleFirstDateCompleted(si.ID);
                
                //iterate people                
                foreach (var ptnd in ptndl)
                {
                    bool bContinue = false;
                    Console.WriteLine(ptnd.Name + ":" + ptnd.sNextDueDate + ", email: " + ptnd.Email + ", Employee ID: " + ptnd.EmployeeID + ", Role: " + ptnd.Role);
                    //just do this for the nurse role
                    if (ptnd.Role != "Nurse")
                        continue;

                    //make sure they are nova net certified
                    if (!ptnd.IsNovaNetTested)
                    {
                        logger.Info("Nova net certification needed for " + ptnd.Name);
                        ptndc2l.Add(ptnd);
                        bContinue = true;
                    }

                    if (ptnd.Email == null)
                    {
                        logger.Info("Email missing for " + ptnd.Name);
                        ptndc3l.Add(ptnd);
                        bContinue = true;
                    }
                    else 
                    {
                        if (ptnd.Email.Trim().Length == 0)
                        {
                            logger.Info("Email missing for " + ptnd.Name);
                            ptndc3l.Add(ptnd);
                            bContinue = true;
                        }
                    }
                    if (si.EmpIDRequired)
                    {
                        if (ptnd.EmployeeID == null)
                        {
                            logger.Info("Employee ID missing for " + ptnd.Name);
                            ptndc4l.Add(ptnd);
                            bContinue = true;
                        }
                        else
                        {
                            if (ptnd.EmployeeID.Trim().Length == 0)
                            {
                                logger.Info("Employee ID missing for " + ptnd.Name);
                                ptndc4l.Add(ptnd);
                                bContinue = true;
                            }
                        }
                    }
                    if (bContinue)
                        continue;

                    //make sure next due date is not null
                    if (ptnd.sNextDueDate != null)
                    {
                        DateTime nd = DateTime.Parse(ptnd.sNextDueDate);
                        TimeSpan ts = nd - DateTime.Now;
                        Console.WriteLine("Window days: " + ts.Days);
                                                
                        if (ts.Days <= 30)
                        {                            
                            //set previous tests to not current (IsCurrent=0)
                            //this allows the user to take the tests again
                            logger.Info("Post tests are due for " + ptnd.Name);
                            
                            int retVal = SetPostTestsCompletedIsCurrent(ptnd.ID);
                            logger.Info("Number of tests set IsCurrent=0: " + retVal);
                            
                            //send email to user                                                           
                            string[] to = new string[] { ptnd.Email };

                            string subject = "Annual Halfpint Post Tests Due";
                            string body = "Your annual halpint post tests are now available at the link below.  Please complete the required tests as soon as possible.";

                            SendHtmlEmail(subject, to, null, body, path, @"<a href='http://halfpintstudy.org/hpProd/PostTests/Initialize'>Halfpint Study Post Tests</a>");
                            
                            //add to list - to be sent to coordinator
                            ptndcl.Add(ptnd);
                        }
                    }
                    else
                    { //do something about null tests

                    }

                } //foreach (var ptnd in ptndl)
                //send the due list to the coordinators
                if (ptndcl.Count > 0)
                {
                    SendCoordinatorsEmailDue(si.ID, ptndcl, path);
                }
                //send the nova net not tested list to the coordinators
                if (ptndc2l.Count > 0)
                {
                    SendCoordinatorsEmailNovaNet(si.ID, ptndc2l, path);
                }
                //send the emails not entered list to the coordinators
                if (ptndc3l.Count > 0)
                {
                    SendCoordinatorsEmailMissingEmail(si.ID, ptndc3l, path);
                }
                //send the employee id not entered list to the coordinators
                if (ptndc4l.Count > 0)
                {
                    SendCoordinatorsEmailMissingEmployeeID(si.ID, ptndc4l, path);
                }
            } //foreach (var si in sites)

            //now update the nova net files
            Console.WriteLine("-------------------------");
            Console.WriteLine("Updating nova net files");
            Console.WriteLine("-------------------------");
            
            //iterate sites
            foreach (var si in sites)
            {
                var lines = GetNovaNetFile(si.Name);
                if (lines == null)
                    continue;               
                
                Console.WriteLine(si.Name);
                var ptndl = GetPostTestPeopleFirstDateCompleted(si.ID);
                
                //iterate people

                var ptndc5l = new List<PostTestNextDue>();
                foreach (var ptnd in ptndl)
                {
                    if (ptnd.EmployeeID == null)
                        continue;
                    if (ptnd.EmployeeID.Trim().Length == 0)
                        continue;
                    if (ptnd.Role != "Nurse")
                        continue;
                    
                    NovaNetColumns line = lines.Find(c => c.EmployeeID == ptnd.EmployeeID );
                    if (line != null)
                    {
                        //make sure they are nova net certified - if not then remove
                        if (!ptnd.IsNovaNetTested)
                        {
                            lines.Remove(line);
                            continue;
                        }

                        DateTime lineDate = DateTime.Parse(line.endDate);
                        DateTime dbDate = DateTime.Parse(ptnd.sNextDueDate);

                        //if the database date is later than the file date 
                        //do an update
                        if (dbDate.CompareTo(lineDate) == 1)
                        {
                            line.endDate = ptnd.NextDueDate.ToString("M/d/yyyy");
                        }
                    }
                    else //this is a new operator
                    {
                        if (!ptnd.IsNovaNetTested)
                            continue;

                        var nnc = new NovaNetColumns();
                        var sep = new[] { ',' };
                        var names = ptnd.Name.Split(sep);
                        nnc.LastName = names[1];
                        nnc.FirstName = names[0];
                        nnc.col3 = "ALL";
                        nnc.col4 = "ALL";
                        nnc.col5 = "StatStrip";
                        nnc.EmployeeID = ptnd.EmployeeID;
                        nnc.col7 = "T";
                        nnc.col8 = "O";
                        nnc.col9 = "Glucose";
                        DateTime start = ptnd.NextDueDate.AddYears(-1);
                        nnc.startDate = start.ToString("M/d/yyyy");
                        nnc.endDate = ptnd.NextDueDate.ToString("M/d/yyyy");
                        lines.Add(nnc);
                    }

                    Console.WriteLine(ptnd.Name + ":" + ptnd.sNextDueDate + ", email: " + ptnd.Email + ", Employee ID: " + ptnd.EmployeeID);
                }
                
                //write lines to new file
                WriteNovaNetFile(si.Name, lines);
                
            } 
            Console.Read();
        }

        public static void SendCoordinatorsEmailDue(int site, List<PostTestNextDue> ptndcl, string path)
        {
            var coordinators = GetUserInRole("Coordinator", site);
            var toEmails = new List<string>();
            foreach (var coord in coordinators)
            {
                toEmails.Add(coord.Email);
            }
            StringBuilder sbBody = new StringBuilder("<p>The following people are due to take their annual post tests.</p>");

            sbBody.Append("<table><tr><th>Name</th><th>Due Date</th><th>Email</th></tr>");
            foreach (var ptnd in ptndcl)
            { 
                sbBody.Append("<tr><td>"+ ptnd.Name + "</td><td>" + ptnd.sNextDueDate + "</td><td>" + ptnd.Email + "</td></tr>");            
            }
            sbBody.Append("</table>");

            SendHtmlEmail("Post Tests - People Due", toEmails.ToArray(), null, sbBody.ToString(), path, @"<a href='http://halfpintstudy.org/hpProd/'>Halfpint Study Website</a>");
        }

        public static void SendCoordinatorsEmailNovaNet(int site, List<PostTestNextDue> ptndcl, string path)
        {
            var coordinators = GetUserInRole("Coordinator", site);
            var toEmails = new List<string>();
            foreach (var coord in coordinators)
            {
                toEmails.Add(coord.Email);
            }
            StringBuilder sbBody = new StringBuilder("<p>The following people have not completed the Nova Net competency test.</p>");

            sbBody.Append("<table><tr><th>Name</th><th>Email</th></tr>");
            string email;
            foreach (var ptnd in ptndcl)
            {
                email = "not entered";
                if (ptnd.Email != null)
                    email = ptnd.Email;
                sbBody.Append("<tr><td>" + ptnd.Name + "</td><td>" + email   + "</td></tr>");
            }
            sbBody.Append("</table>");

            SendHtmlEmail("Post Tests - People Due", toEmails.ToArray(), null, sbBody.ToString(), path, @"<a href='http://halfpintstudy.org/hpProd/'>Halfpint Study Website</a>");
        }

        public static void SendCoordinatorsEmailMissingEmail(int site, List<PostTestNextDue> ptndcl, string path)
        {
            var coordinators = GetUserInRole("Coordinator", site);
            var toEmails = new List<string>();
            foreach (var coord in coordinators)
            {
                toEmails.Add(coord.Email);
            }
            StringBuilder sbBody = new StringBuilder("<p>The following people need to have their email address entered into the staff table.</p>");

            sbBody.Append("<table><tr><th>Name</th></tr>");
            foreach (var ptnd in ptndcl)
            {
                sbBody.Append("<tr><td>" + ptnd.Name + "</td></tr>");
            }
            sbBody.Append("</table>");

            SendHtmlEmail("Post Tests - Staff - Employee Emails Missing", toEmails.ToArray(), null, sbBody.ToString(), path, @"<a href='http://halfpintstudy.org/hpProd/'>Halfpint Study Website</a>");
        }

        public static void SendCoordinatorsEmailMissingEmployeeID(int site, List<PostTestNextDue> ptndcl, string path)
        {
            var coordinators = GetUserInRole("Coordinator", site);
            var toEmails = new List<string>();
            foreach (var coord in coordinators)
            {
                toEmails.Add(coord.Email);
            }
            StringBuilder sbBody = new StringBuilder("<p>The following people need to have their employee ID entered into the staff table.</p>");

            sbBody.Append("<table><tr><th>Name</th><th>Email</th></tr>");
            string email;
            foreach (var ptnd in ptndcl)
            {
                email = "not entered";
                if (ptnd.Email != null)
                    email = ptnd.Email;
                sbBody.Append("<tr><td>" + ptnd.Name + "</td><td>" + email + "</td></tr>");
            }
            sbBody.Append("</table>");

            SendHtmlEmail("Post Tests - Staff - Employee ID's Missing", toEmails.ToArray(), null, sbBody.ToString(), path, @"<a href='http://halfpintstudy.org/hpProd/'>Halfpint Study Website</a>");
        }

        public static List<MembershipUser> GetUserInRole(string role, int site)
        {
            var memUsers = new List<MembershipUser>();
            string[] users = Roles.GetUsersInRole(role);

            String strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            using (SqlConnection conn = new SqlConnection(strConn))
            {
                try
                {
                    SqlCommand cmd = new SqlCommand("", conn);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.CommandText = ("GetSiteUsers");
                    SqlParameter param = new SqlParameter("@siteID", site);
                    cmd.Parameters.Add(param);

                    conn.Open();
                    SqlDataReader rdr = cmd.ExecuteReader();

                    int pos = 0;
                    string userName = "";
                    while (rdr.Read())
                    {
                        pos = rdr.GetOrdinal("UserName");
                        userName = rdr.GetString(pos);
                        foreach (var u in users)
                        {
                            if (u == userName)
                            {
                                memUsers.Add(Membership.GetUser(u));
                            }
                        }
                    }
                    rdr.Close();
                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                }
            }
            return memUsers;
        }

        static int SetPostTestsCompletedIsCurrent(int id)
        {            
            String strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
             using (SqlConnection conn = new SqlConnection(strConn))
            {
                try
                {
                    SqlCommand cmd = new SqlCommand("", conn);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.CommandText = "SetStaffPostTestsCompletedIsCurrent";

                    SqlParameter param = new SqlParameter("@id", id);
                    cmd.Parameters.Add(param);

                    conn.Open();
                    return cmd.ExecuteNonQuery();
                    
                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                    return -1;
                }
            }
            
        }

        static List<SiteInfo> GetSites()
        {
            var sil = new List<SiteInfo>();

            String strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();

            var si = new SiteInfo();
            using (SqlConnection conn = new SqlConnection(strConn))
            {
                try
                {
                    SqlCommand cmd = new SqlCommand("", conn);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.CommandText = "GetSites";
                    
                    conn.Open();
                    int pos = 0;
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        si = new SiteInfo();
                        pos = rdr.GetOrdinal("ID");
                        si.ID = rdr.GetInt32(pos);
                        pos = rdr.GetOrdinal("Name");
                        si.Name = rdr.GetString(pos);
                        pos = rdr.GetOrdinal("SiteID");
                        si.SiteID = rdr.GetString(pos);
                        pos = rdr.GetOrdinal("EmpIDRequired");
                        si.EmpIDRequired = rdr.GetBoolean(pos);
                        sil.Add(si);
                    }
                    rdr.Close();
                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                }
            }
            return sil;
        }
                
        static List<PostTestNextDue> GetPostTestPeopleFirstDateCompleted(int siteID)
        {
            var ptndl = new List<PostTestNextDue>();
            var ptnd = new PostTestNextDue();

            String strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            using (SqlConnection conn = new SqlConnection(strConn))
            {
                try
                {
                    SqlCommand cmd = new SqlCommand("", conn);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.CommandText = ("GetStaffPostTestsFirstDateCompletedBySite");
                    SqlParameter param = new SqlParameter("@siteID", siteID);
                    cmd.Parameters.Add(param);

                    conn.Open();
                    SqlDataReader rdr = cmd.ExecuteReader();
                    int pos = 0;

                    while (rdr.Read())
                    {
                        ptnd = new PostTestNextDue();

                        pos = rdr.GetOrdinal("ID");
                        ptnd.ID = rdr.GetInt32(pos);

                        pos = rdr.GetOrdinal("Name");
                        ptnd.Name = rdr.GetString(pos);

                        pos = rdr.GetOrdinal("MinDate");
                        if (!rdr.IsDBNull(pos))
                        {
                            ptnd.NextDueDate = rdr.GetDateTime(pos).AddYears(1);
                            ptnd.sNextDueDate = ptnd.NextDueDate.ToString("MM/dd/yyyy");
                        }

                        pos = rdr.GetOrdinal("Email");
                        if (!rdr.IsDBNull(pos))
                        {
                            ptnd.Email = rdr.GetString(pos);
                        }

                        pos = rdr.GetOrdinal("EmployeeID");
                        if (!rdr.IsDBNull(pos))
                        {
                            ptnd.EmployeeID = rdr.GetString(pos);
                        }

                        pos = rdr.GetOrdinal("NovaStatStrip");
                        if (!rdr.IsDBNull(pos))
                        {
                            ptnd.IsNovaNetTested = rdr.GetBoolean(pos);
                        }

                        pos = rdr.GetOrdinal("Role");
                        if (!rdr.IsDBNull(pos))
                        {
                            ptnd.Role = rdr.GetString(pos);
                        }
                        ptndl.Add(ptnd);
                    }
                    rdr.Close();
                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                    return null;
                }
            }
            return ptndl;
        }

        public static void SendHtmlEmail(string subject, string[] toAddress, string[] ccAddress, string body, string appPath, string url, string bodyHeader = "")
        {
            
            MailMessage mm = new MailMessage();
            mm.Subject = subject;
            //mm.IsBodyHtml = true;
            mm.Body = body;
            string path = Path.Combine(appPath, "mailLogo.jpg");
            LinkedResource mailLogo = new LinkedResource(path);

            StringBuilder sb = new StringBuilder("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0 Transitional//EN\">");
            sb.Append("<html>");
            sb.Append("<head>");
            //sb.Append("<style type='text/css'>");
            //sb.Append("body {margin:50px 0px; padding:0px; text-align:center; }");
            //sb.Append("#Content {width:500px; margin:0px auto; text-align:left; padding:15px; border:1px dashed #333; background-color:#eee;}");


            //sb.Append("</style");
            sb.Append("</head>");
            sb.Append("<body style='text-align:center;'>");
            sb.Append("<img style='width:200px;' alt='' hspace=0 src='cid:mailLogoID' align=baseline />");
            if (bodyHeader.Length > 0)
            {
                sb.Append(bodyHeader);
            }

            sb.Append("<div style='text-align:left;margin-left:30px;'>");
            sb.Append("<table style='margin-left:0px;'>");
            sb.Append(body);
            sb.Append("</table>");
            sb.Append("<br/><br/>" + url);
            sb.Append("</div>");
            sb.Append("</body>");
            sb.Append("</html>");
                       
            AlternateView av = AlternateView.CreateAlternateViewFromString(sb.ToString(), null, "text/html");

            mailLogo.ContentId = "mailLogoID";
            av.LinkedResources.Add(mailLogo);
            mm.AlternateViews.Add(av);

            foreach (string s in toAddress)
                mm.To.Add(s);
            if (ccAddress != null)
            {
                foreach (string s in ccAddress)
                    mm.CC.Add(s);
            }

            SmtpClient smtp = new SmtpClient();
            smtp.Send(mm);
        }

        static void WriteNovaNetFile(string site, List<NovaNetColumns> lines)
        {
            string filePart = " StatStrip Nurse Training List.csv";
            string file = site + filePart;

            //write lines to new file
            string fileName = Path.Combine("C:\\Halfpint\\NovaNet\\OperatorsList", "test.csv");
            
            StreamWriter sw = new StreamWriter(fileName, false);


            sw.WriteLine("NovaNet Operator Import Data,version 2.0,,,,,,,,,");
            foreach (var line in lines)
            {
                sw.Write(line.LastName + ",");
                sw.Write(line.FirstName + ",");
                sw.Write(line.col3 + ",");
                sw.Write(line.col4 + ",");
                sw.Write(line.col5 + ",");
                sw.Write(line.EmployeeID + ",");
                sw.Write(line.col7 + ",");
                sw.Write(line.col8 + ",");
                sw.Write(line.col9 + ",");
                sw.Write(line.startDate + ",");
                sw.Write(line.endDate);
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        static List<NovaNetColumns> GetNovaNetFile(string site)
        {
            DirectoryInfo di = new DirectoryInfo("C:\\Halfpint\\NovaNet\\OperatorsList");
            var files = di.EnumerateFiles();
            string fileName = "";

            foreach (var file in files)
            {
                if (file.Name.StartsWith(site))
                {
                    fileName = file.Name;
                    break;
                }
            }
            if (fileName.Length == 0)
                return null;

            fileName = Path.Combine("C:\\Halfpint\\NovaNet\\OperatorsList", fileName);
            List<NovaNetColumns> lines = new List<NovaNetColumns>();
            char[] delimiters = new char[] { ',' };
            using (StreamReader reader = new StreamReader(fileName))
            {
                var cols = new NovaNetColumns();
                int count = 1;
                while (true)
                {
                    string line = reader.ReadLine();
                    if (line == null)
                    {
                        break;
                    }
                    
                    if (count == 1)
                    {                        
                        count = 2;
                        continue;
                    }
                    else
                    {
                        cols = new NovaNetColumns();
                        string[] parts = line.Split(delimiters);
                        string empID = "";
                    
                        cols.LastName = parts[0];
                        cols.FirstName = parts[1];
                        cols.col3 = parts[2];
                        cols.col4 = parts[3];
                        cols.col5 = parts[4];

                        empID = parts[5];
                        switch(site)
                        {
                            case "CHB":
                                if (empID.Length < 6)
                                {
                                    var add = 6 - empID.Length;
                                    for (int i = 0; i < add; i++)
                                        empID = "0" + empID;                                    
                                }
                                    
                                break;
                        }
                        cols.EmployeeID = empID;
                        cols.col7 = parts[6];
                        cols.col8 = parts[7];
                        cols.col9 = parts[8];
                        cols.startDate = parts[9];
                        cols.endDate = parts[10];
                    }

                    lines.Add(cols);
                    // Console.WriteLine("{0} field(s)", parts.Length);
                }
            }
            return lines;
        }

        
    }

    
    public class SiteInfo
    {
        public int ID { get; set; }
        public string SiteID { get; set; }
        public string Name { get; set; }
        public bool EmpIDRequired { get; set; }
    }

    public class PostTestNextDue
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public DateTime NextDueDate { get; set; }
        public string sNextDueDate { get; set; }
        public string Email { get; set; }
        public string EmployeeID { get; set; }
        public bool IsNovaNetTested { get; set; }
        public string Role { get; set; }
    }

    public class NovaNetColumns
    {
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string col3 { get; set; }
        public string col4 { get; set; }
        public string col5 { get; set; }
        public string EmployeeID { get; set; }
        public string col7 { get; set; }
        public string col8 { get; set; }
        public string col9 { get; set; }
        public string startDate { get; set; }
        public string endDate { get; set; }
    }

}
