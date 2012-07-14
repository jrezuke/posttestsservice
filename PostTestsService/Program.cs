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
            var ptndcl = new List<PostTestNextDue>();

            //iterate sites
            foreach (var si in sites)
            {
                Console.WriteLine(si.Name);
                
                //Get the next date due for people - this works on tests
                //that are current (IsCurrent=1)
                var ptndl = GetPostTestPeopleFirstDateCompleted(si.ID);
                
                //iterate people
                
                foreach (var ptnd in ptndl)
                {
                    Console.WriteLine(ptnd.Name + ":" + ptnd.sNextDueDate + ", email: " + ptnd.Email + ", Employee ID: " + ptnd.EmployeeID);
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
                            
                            int retVal = SetPostTestsCompletedIsCurrent(ptnd.Name);
                            logger.Info("Number of tests set IsCurrent=0: " + retVal);
                            
                            //send email to user
                            if (ptnd.Email != null)                                
                            {
                                if (ptnd.Email.Trim().Length >0 ) 
                                {
                                    string[] to = new string[] { ptnd.Email };
                                    
                                    string subject = "Annual Halfpint Post Tests Due";
                                    string body = "Your annual halpint post tests are now available at the link below.  Please complete the required tests as soon as possible.";

                                    SendHtmlEmail(subject, to, null, body, path, @"<a href='http://halfpintstudy.org/hpProd/PostTests/Initialize'>Halfpint Study Post Tests</a>");

                                }

                            }

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
                    var coordinators = GetUserInRole("Coordinator", si.ID);
                    var toEmails = new List<string>();
                    foreach (var coord in coordinators)
                    {
                        toEmails.Add(coord.Email);
                    }
                    
                }

            } //foreach (var si in sites)
            Console.Read();
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

        static int SetPostTestsCompletedIsCurrent(string name)
        {            
            String strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
             using (SqlConnection conn = new SqlConnection(strConn))
            {
                try
                {
                    SqlCommand cmd = new SqlCommand("", conn);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.CommandText = "SetPostTestsCompletedIsCurrent";

                    SqlParameter param = new SqlParameter("@name", name);
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
                    cmd.CommandText = ("GetPostTestPeopleFirstDateCompleted");
                    SqlParameter param = new SqlParameter("@siteID", siteID);
                    cmd.Parameters.Add(param);

                    conn.Open();
                    SqlDataReader rdr = cmd.ExecuteReader();
                    int pos = 0;

                    while (rdr.Read())
                    {
                        ptnd = new PostTestNextDue();

                        pos = rdr.GetOrdinal("Name");
                        ptnd.Name = rdr.GetString(pos);

                        pos = rdr.GetOrdinal("MinDate");
                        if (!rdr.IsDBNull(pos))
                        {
                            ptnd.NextDueDate = rdr.GetDateTime(pos).AddYears(1);
                            ptnd.sNextDueDate = ptnd.NextDueDate.Value.ToString("MM/dd/yyyy");
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

            //string sAv = "<img alt='' hspace=0 src='cid:mailLogoID' align=baseline /><br/>";
            //sAv += body;


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
    }

    
    public class SiteInfo
    {
        public int ID { get; set; }
        public string SiteID { get; set; }
        public string Name { get; set; }
    }

    public class PostTestNextDue
    {
        public string Name { get; set; }
        public DateTime? NextDueDate { get; set; }
        public string sNextDueDate { get; set; }
        public string Email { get; set; }
        public string EmployeeID { get; set; }
    }


}
