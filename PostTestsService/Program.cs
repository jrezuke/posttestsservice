using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NLog;
using System.Configuration;
using System.Net.Mail;
using System.IO;
using System.Web.Security;

namespace PostTestsService
{
    class Program
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        
        static void Main(string[] args)
        {       
            Logger.Info("Starting PostTests Service");
            
            string path = AppDomain.CurrentDomain.BaseDirectory;
            //this will be a parsed list for the 2nd run
            var postTestNextDues2 = new List<PostTestNextDue>();
           
            //get sites 
            var sites = GetSites();
            var siteLists = new List<SiteEmailLists>();
            
            //iterate sites
            foreach (var si in sites)
            {
                var siteEmailList = new SiteEmailLists
                                        {
                                            SiteId = si.Id,
                                            DueList = new List<PostTestNextDue>(),
                                            CompetencyMissingList = new List<PostTestNextDue>(),
                                            EmailMissingList = new List<PostTestNextDue>(),
                                            EmployeeIdMissingList = new List<PostTestNextDue>()
                                        };

                siteLists.Add(siteEmailList);

                Console.WriteLine(si.Name);
                Logger.Info("For Site:" + si.Name + " - " + si.SiteId);
                Logger.Debug("For Site:" + si.Name + " - " + si.SiteId);
                
                //Get the next date due for people - this works on tests
                //that are current (IsCurrent=1)
                var postTestNextDueList = GetAllStaffPostTestsCompleted(si.Id); //GetPostTestPeopleFirstDateCompleted(si.Id);
                
                //iterate people                
                foreach (var postTestNextDue in postTestNextDueList)
                {
                    bool bContinue = false;
                    Console.WriteLine(postTestNextDue.Name + ":" + postTestNextDue.SNextDueDate + ", email: " + postTestNextDue.Email + ", Employee ID: " + postTestNextDue.EmployeeId + ", Role: " + postTestNextDue.Role);
                    //just do this for the nurse role
                    if (postTestNextDue.Role != "Nurse")
                        continue;

                    //make sure they are nova net certified
                    if ((!postTestNextDue.IsNovaNetTested) || (!postTestNextDue.IsVampTested))
                    {
                        Logger.Info("Competency needed for " + postTestNextDue.Name);
                        siteEmailList.CompetencyMissingList.Add(postTestNextDue);
                        bContinue = true;
                    }

                    if (postTestNextDue.Email == null)
                    {
                        Logger.Info("Email missing for " + postTestNextDue.Name);
                        siteEmailList.EmailMissingList.Add(postTestNextDue);
                        bContinue = true;
                    }
                    else 
                    {
                        if (postTestNextDue.Email.Trim().Length == 0)
                        {
                            Logger.Info("Email missing for " + postTestNextDue.Name);
                            siteEmailList.EmailMissingList.Add(postTestNextDue);
                            bContinue = true;
                        }
                    }
                    if (si.EmpIdRequired)
                    {
                        if (postTestNextDue.EmployeeId == null)
                        {
                            Logger.Info("Employee ID missing for " + postTestNextDue.Name);
                            siteEmailList.EmployeeIdMissingList.Add(postTestNextDue);
                            bContinue = true;
                        }
                        else
                        {
                            if (postTestNextDue.EmployeeId.Trim().Length == 0)
                            {
                                Logger.Info("Employee ID missing for " + postTestNextDue.Name);
                                siteEmailList.EmployeeIdMissingList.Add(postTestNextDue);
                                bContinue = true;
                            }
                        }
                    }
                    if (bContinue)
                        continue;

                    //create list for 2nd run
                    postTestNextDues2.Add(postTestNextDue);

                    //if (postTestNextDue.SNextDueDate == null) then either new staff or all tests completed are not current 
                    
                    if ( postTestNextDue.SNextDueDate != null)
                    {
                        var nd = DateTime.Parse(postTestNextDue.SNextDueDate);
                        var ts = nd - DateTime.Now;
                        Console.WriteLine("Window days: " + ts.Days);
                        if (ts.Days > 30) continue;
                    
                    }
                    
                    Logger.Info("Post tests are due for " + postTestNextDue.Name);

                    //set previous tests to not current (IsCurrent=0)
                    //this allows the user to take the tests again
                    int retVal = SetPostTestsCompletedIsCurrent(postTestNextDue.Id);
                    Logger.Info("Number of tests set IsCurrent=0: " + retVal);
                            
                    //send email to user                                                           
                    var to = new[] { postTestNextDue.Email };

                    const string subject = "Annual Halfpint Post Tests Due";
                    const string body = "Your annual halpint post tests are now available at the link below.  Please complete the required tests as soon as possible.";

                    SendHtmlEmail(subject, to, null, body, path, @"<a href='http://halfpintstudy.org/hpProd/PostTests/Initialize'>Halfpint Study Post Tests</a>");
                            
                    //add to list - to be sent to coordinator
                    siteEmailList.DueList.Add(postTestNextDue);
                } //foreach (var ptnd in ptndl)
                
            } //foreach (var si in sites)
            
            //now update the nova net files
            Console.WriteLine("-------------------------");
            Console.WriteLine("Updating nova net files");
            Console.WriteLine("-------------------------");
            
            //iterate sites
            foreach (var si in sites)
            {
                var siteEmailList = siteLists.Find(x => x.SiteId == si.Id);

                siteEmailList.StaffAddedList = new List<PostTestNextDue>();
                siteEmailList.StaffRemovedList = new List<PostTestNextDue>();

                //skip for sites not needed
                if (!si.EmpIdRequired)
                    continue;

                Console.WriteLine(si.Name);
                Logger.Info("For Site:" + si.Name + " - " + si.SiteId);

                var lines = GetNovaNetFile(si.Name);
                if (lines == null)
                {
                    Console.WriteLine("No current nova net list for site:" + si.Name);
                    Logger.Info("No current nova net list for site:" + si.Name);
                    Console.WriteLine("created new list");
                    Logger.Info("created new list");
                    
                    //create the new list
                    lines = new List<NovaNetColumns>();
                }

                //var ptndl = GetPostTestPeopleFirstDateCompleted(si.Id);
                
                //iterate people
                foreach (var ptnd in postTestNextDues2)
                {
                    //this is now handled in postTestNextDues2
                    //if (ptnd.EmployeeId == null)
                    //    continue;
                    //if (ptnd.EmployeeId.Trim().Length == 0)
                    //    continue;
                    //if (ptnd.Role != "Nurse")
                    //    continue;
                    
                    NovaNetColumns line = lines.Find(c => c.EmployeeId == ptnd.EmployeeId );
                    if (line != null)
                    {
                        //make sure they are certified - if not then remove
                        //this is now handled in postTestNextDues2  
                        //if ((!ptnd.IsNovaNetTested) || (!ptnd.IsVampTested))
                        //{
                        //    lines.Remove(line);
                        //    siteEmailList.StaffRemovedList.Add(ptnd);
                        //    continue;
                        //}

                        line.Found = true;
                        var lineDate = DateTime.Parse(line.EndDate);
                        var endDate = DateTime.Now.AddYears(1);
                        if(! string.IsNullOrEmpty(ptnd.SNextDueDate))
                            endDate = DateTime.Parse(ptnd.SNextDueDate);

                        //update the line end date
                        line.EndDate = endDate.ToString("M/d/yyyy");
                    }
                    else //this is a new operator
                    {
                        //make sure they are certified - if not then don't add
                        //if ((!ptnd.IsNovaNetTested) || (!ptnd.IsVampTested))
                        //    continue;

                        //email coord
                        siteEmailList.StaffAddedList.Add(ptnd);
                        var nnc = new NovaNetColumns();
                        var sep = new[] { ',' };
                        var names = ptnd.Name.Split(sep);
                        nnc.LastName = names[1];
                        nnc.FirstName = names[0];
                        nnc.Col3 = "ALL";
                        nnc.Col4 = "ALL";
                        nnc.Col5 = "StatStrip";
                        nnc.EmployeeId = ptnd.EmployeeId;
                        nnc.Col7 = "T";
                        nnc.Col8 = "O";
                        nnc.Col9 = "Glucose";

                        DateTime startDate = DateTime.Now;
                        DateTime endDate = DateTime.Now.AddYears(1);
                        
                        nnc.StartDate = startDate.ToString("M/d/yyyy");
                        nnc.EndDate = endDate.ToString("M/d/yyyy");
                        lines.Add(nnc);
                    }

                    Console.WriteLine(ptnd.Name + ":" + ptnd.SNextDueDate + ", email: " + ptnd.Email + ", Employee ID: " + ptnd.EmployeeId);
                }
                
                //remove the nn lines that were not found


                //write lines to new file
                WriteNovaNetFile(lines);

            }//foreach (var si in sites) - write file

            foreach (var si in sites)
            {
                var siteEmailList = siteLists.Find(x => x.SiteId == si.Id);

                siteEmailList.StaffTestsNotCompletedList = GetStaffPostTestCompleted(si.Id).ToList();
                
                SendCoordinatorsEmail(si.Id, siteEmailList, path);
            }//foreach (var si in sites) - tests not completed
           
            Console.Read();
        }

        
        public static void SendCoordinatorsEmail(int site, SiteEmailLists siteEmailLists, string path)
        {
            var coordinators = GetUserInRole("Coordinator", site);
            var sbBody = new StringBuilder("");
            const string newLine = "<br/>";
            
            sbBody.Append(newLine);
            if (siteEmailLists.DueList.Count == 0)
                sbBody.Append("<h3>There are no nurses due to take their annual post tests.</h3>");
            else
            {
                var dueSortedList = siteEmailLists.DueList.OrderBy(x => x.NextDueDate).ToList();
                sbBody.Append("<h3>The following nurses are due to take their annual post tests.</h3>");

                sbBody.Append("<table style='border-collapse:collapse;' cellpadding='5' border='1'><tr style='background-color:87CEEB'><th>Name</th><th>Due Date</th><th>Email</th></tr>");
                foreach (var ptnd in dueSortedList)
                {
                    sbBody.Append("<tr><td>" + ptnd.Name + "</td><td>" + ptnd.SNextDueDate + "</td><td>" + ptnd.Email +
                                  "</td></tr>");
                }
                sbBody.Append("</table>");
            }

            if (siteEmailLists.CompetencyMissingList.Count == 0)
                sbBody.Append("<h3>All nurses have completed threir competency tests.</h3>");
            else
            {
                var competencyMissingSortedList = siteEmailLists.CompetencyMissingList.OrderBy(x => x.Name).ToList();
                sbBody.Append("<h3>The following people have not completed a competency test.</h3>");

                sbBody.Append("<table style='border-collapse:collapse;' cellpadding='5' border='1'><tr style='background-color:87CEEB'><th>Name</th><th>Tests Not Completed</th><th>Email</th></tr>");
                foreach (var ptnd in competencyMissingSortedList)
                {
                    var email = "not entered";
                    if (ptnd.Email != null)
                        email = ptnd.Email;
                    
                    var test = "";
                    if (!ptnd.IsNovaNetTested)
                        test = "Nova Net";
                    if (!ptnd.IsVampTested)
                    {
                        if (test.Length > 0)
                            test += " and ";
                        test += "Vamp Jr";
                    }

                    sbBody.Append("<tr><td>" + ptnd.Name + "</td><td>" + test + "</td><td>" + email + "</td></tr>");
                }
                sbBody.Append("</table>");
            }

            if (siteEmailLists.EmailMissingList.Count > 0)
            {
                var emailMissingSortedList = siteEmailLists.EmailMissingList.OrderBy(x => x.Name).ToList();
                sbBody.Append("<h3>The following nurses need to have their email address entered into the staff table.</h3>");

                sbBody.Append("<div><table style='border-collapse:collapse;' cellpadding='5' border='1'><tr style='background-color:87CEEB'><th>Name</th></tr>");
                foreach (var ptnd in emailMissingSortedList)
                {
                    sbBody.Append("<tr><td>" + ptnd.Name + "</td></tr>");
                }
                sbBody.Append("</table></div>");
            }

            if (siteEmailLists.EmployeeIdMissingList.Count > 0)
            {
                var employeeIdMissingSortedList = siteEmailLists.EmployeeIdMissingList.OrderBy(x => x.Name).ToList();
                sbBody.Append("<h3>The following nurses need to have their employee ID entered into the staff table.</h3>");

                sbBody.Append("<div><table style='border-collapse:collapse;' cellpadding='5' border='1'><tr style='background-color:87CEEB'><th>Name</th><th>Email</th></tr>");
                foreach (var ptnd in employeeIdMissingSortedList)
                {
                    var email = "not entered";
                    if (ptnd.Email != null)
                        email = ptnd.Email;
                    sbBody.Append("<tr><td>" + ptnd.Name + "</td><td>" + email + "</td></tr>");
                }
                sbBody.Append("</table></div>");
            }

            if(siteEmailLists.StaffTestsNotCompletedList.Count>0)
            {
                sbBody.Append("<h3>The following nurses have not completed all post tests.</h3>");

                sbBody.Append("<div><table style='border-collapse:collapse;' cellpadding='5' border='1';><tr style='background-color:87CEEB'><th>Name</th><th>Email</th><th>Tests Not Completed</th></tr>");
                foreach (var tncl in siteEmailLists.StaffTestsNotCompletedList)
                {
                    //sbBody.Append("<div>");
                    var email = "not entered";
                    if (tncl.Email != null)
                        email = tncl.Email;
                    
                    sbBody.Append("<tr><td>" + tncl.StaffName + "</td><td>" + email + "</td><td>");
                    
                    foreach (var test in tncl.TestsNotCompleted)
                    {
                        sbBody.Append(test + newLine);
                    }
                    sbBody.Append("</td></tr>");
                    
                }
                sbBody.Append("</table></div>");
                
            }
            
            SendHtmlEmail("Post Tests Notifications", coordinators.Select(coord => coord.Email).ToArray(), null, sbBody.ToString(), path, @"<a href='http://halfpintstudy.org/hpProd/'>Halfpint Study Website</a>");
        }

        public static void SendHtmlEmail(string subject, string[] toAddress, string[] ccAddress, string bodyContent, string appPath, string url, string bodyHeader = "")
        {

            if (toAddress.Length == 0)
                return;
            var mm = new MailMessage { Subject = subject, Body = bodyContent };
            //mm.IsBodyHtml = true;
            var path = Path.Combine(appPath, "mailLogo.jpg");
            var mailLogo = new LinkedResource(path);

            var sb = new StringBuilder("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0 Transitional//EN\">");
            sb.Append("<html>");
            sb.Append("<head>");
            
            //sb.Append("<style type='text/css'>");
            //sb.Append("td {padding:10px; }");
            //sb.Append("#Content {width:500px; margin:0px auto; text-align:left; padding:15px; border:1px dashed #333; background-color:#eee;}");
            //sb.Append("</style");
            
            sb.Append("</head>");
            sb.Append("<body style='text-align:left;'>");
            sb.Append("<img style='width:200px;' alt='' hspace=0 src='cid:mailLogoID' align=baseline />");
            if (bodyHeader.Length > 0)
            {
                sb.Append(bodyHeader);
            }

            sb.Append("<div style='text-align:left;margin-left:30px;width:100%'>");
            sb.Append("<table style='margin-left:0px;'>");
            sb.Append(bodyContent);
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

            Console.WriteLine("Send Email");
            Console.WriteLine("Subject");

            var smtp = new SmtpClient();
            //smtp.Send(mm);
        }

        public static List<MembershipUser> GetUserInRole(string role, int site)
        {
            var memUsers = new List<MembershipUser>();
            string[] users = Roles.GetUsersInRole(role);

            String strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            using (var conn = new SqlConnection(strConn))
            {
                try
                {
                    var cmd = new SqlCommand("", conn)
                                  {
                                      CommandType = System.Data.CommandType.StoredProcedure,
                                      CommandText = ("GetSiteUsers")
                                  };
                    var param = new SqlParameter("@siteID", site);
                    cmd.Parameters.Add(param);

                    conn.Open();
                    var rdr = cmd.ExecuteReader();

                    while (rdr.Read())
                    {
                        var pos = rdr.GetOrdinal("UserName");
                        var userName = rdr.GetString(pos);
                        memUsers.AddRange(from u in users where u == userName select Membership.GetUser(u));
                    }
                    rdr.Close();
                }
                catch (Exception ex)
                {
                    Logger.Error(ex);
                }
            }
            return memUsers;
        }

        static int SetPostTestsCompletedIsCurrent(int id)
        {            
            String strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
             using (var conn = new SqlConnection(strConn))
            {
                try
                {
                    var cmd = new SqlCommand("", conn)
                                  {
                                      CommandType = System.Data.CommandType.StoredProcedure,
                                      CommandText = "SetStaffPostTestsCompletedIsCurrent"
                                  };

                    var param = new SqlParameter("@id", id);
                    cmd.Parameters.Add(param);

                    conn.Open();
                    return cmd.ExecuteNonQuery();
                    
                }
                catch (Exception ex)
                {
                    Logger.Error(ex);
                    return -1;
                }
            }
            
        }

        static List<SiteInfo> GetSites()
        {
            var sil = new List<SiteInfo>();

            String strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();

            using (var conn = new SqlConnection(strConn))
            {
                try
                {
                    var cmd = new SqlCommand("", conn)
                                  {CommandType = System.Data.CommandType.StoredProcedure, CommandText = "GetSitesActive"};

                    conn.Open();
                    var rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var si = new SiteInfo();
                        var pos = rdr.GetOrdinal("ID");
                        si.Id = rdr.GetInt32(pos);
                        pos = rdr.GetOrdinal("Name");
                        si.Name = rdr.GetString(pos);
                        pos = rdr.GetOrdinal("SiteID");
                        si.SiteId = rdr.GetString(pos);
                        pos = rdr.GetOrdinal("EmpIDRequired");
                        si.EmpIdRequired = rdr.GetBoolean(pos);
                        sil.Add(si);
                    }
                    rdr.Close();
                }
                catch (Exception ex)
                {
                    Logger.Error(ex);
                }
            }
            return sil;
        }

        static IEnumerable<StaffTestsNotCompletedList> GetStaffPostTestCompleted(int siteId)
        {
            var tncll = new List<StaffTestsNotCompletedList>();
            bool anyTestsNotCompleted = false;
            bool isFirstOne = true;

            String strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            using (var conn = new SqlConnection(strConn))
            {
                try
                {
                    var cmd = new SqlCommand("", conn)
                                  {
                                      CommandType = System.Data.CommandType.StoredProcedure,
                                      CommandText = ("GetStaffPostTestsCompleted")
                                  };
                    var param = new SqlParameter("@siteID", siteId);
                    cmd.Parameters.Add(param);

                    conn.Open();

                    var rdr = cmd.ExecuteReader();
                    
                    var staffId = 0;
                    var tncl = new StaffTestsNotCompletedList();
                    while (rdr.Read())
                    {
                        int pos = 0;

                        pos = rdr.GetOrdinal("ID");
                        int id = rdr.GetInt32(pos);

                        if(staffId != id)
                        {
                            if (isFirstOne)
                                isFirstOne = false;
                            else
                            {
                                if(tncl.TestsNotCompleted.Count> 0)
                                    tncll.Add(tncl);        
                            }
                            tncl = new StaffTestsNotCompletedList {StaffId = id};
                            staffId = id;

                            pos = rdr.GetOrdinal("Email");
                            if(!rdr.IsDBNull(pos))
                                tncl.Email = rdr.GetString(pos);

                            string lastName = "";
                            pos = rdr.GetOrdinal("LastName");
                            if(!rdr.IsDBNull(pos))
                                lastName = rdr.GetString(pos);
                            string firstName = "";
                            pos = rdr.GetOrdinal("FirstName");
                            if (!rdr.IsDBNull(pos))
                                firstName = rdr.GetString(pos);
                            tncl.StaffName = lastName + ", " + firstName;
                        }
                            
                        
                        pos = rdr.GetOrdinal("IsCurrent");
                        if (rdr.IsDBNull(pos))
                            continue;

                        bool isCurrent = rdr.GetBoolean(pos);

                        if(isCurrent)
                        {
                            pos = rdr.GetOrdinal("TestName");
                            string test = rdr.GetString(pos);
                            tncl.TestsNotCompleted.Remove(test);  
                        }

                    }
                    rdr.Close();
                }
                catch (Exception ex)
                {
                    Logger.Error(ex);
                    return null;
                }
            }
            return tncll;
        }

        private static IEnumerable<PostTestNextDue> GetAllStaffPostTestsCompleted(int siteId)
        {
            var ptndl = new List<PostTestNextDue>();

            var strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            using (var conn = new SqlConnection(strConn))
            {
                try
                {
                    var cmd = new SqlCommand("", conn)
                                  {
                                      CommandType = System.Data.CommandType.StoredProcedure,
                                      CommandText = ("GetStaffAllPostTestsDateCompletedBySite")
                                  };
                    var param = new SqlParameter("@siteID", siteId);
                    cmd.Parameters.Add(param);

                    conn.Open();

                    var curId = 0;
                    string sMinDate = null;
                    string name = null;
                    string email = null;
                    string empId = null;

                    var minDueDate = DateTime.Parse("01/01/2100");
                    var nextDueDate = new DateTime();
                    var bHasCompleteDate = false;
                    var bVampCompleted = false;
                    var bNovaNetCompleted = false;
                    
                    var rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var pos = rdr.GetOrdinal("Role");
                        var role = rdr.GetString(pos);
                        var bIsCurrent = false;

                        if(role != "Nurse")
                            continue;
                        
                        pos = rdr.GetOrdinal("ID");
                        var iD = rdr.GetInt32(pos);
                        
                        if (iD != curId)
                        {
                            //add to list if current id changes
                            if (curId != 0)
                            {
                                var ptnd = new PostTestNextDue();
                                ptnd.Id = curId;
                                ptnd.Name = name;
                                ptnd.Role = role;
                                ptnd.Email = email;
                                ptnd.EmployeeId = empId;
                                ptnd.IsNovaNetTested = bNovaNetCompleted;
                                ptnd.IsVampTested = bVampCompleted;
                                if (bHasCompleteDate)
                                {
                                    ptnd.NextDueDate = minDueDate.AddYears(1);
                                    ptnd.SNextDueDate = ptnd.NextDueDate.ToString("MM/dd/yyyy");
                                }
                                ptndl.Add(ptnd);
                            }
                            curId = iD;
                            minDueDate = DateTime.Parse("01/01/2100");
                            nextDueDate = new DateTime();
                            bHasCompleteDate = false;
                            bVampCompleted = false;
                            bNovaNetCompleted = false;
                            name = null;
                            email = null;
                            empId = null;
                        }
                        
                        pos = rdr.GetOrdinal("Name");
                        name = rdr.GetString(pos);

                        pos = rdr.GetOrdinal("IsCurrent");
                        if (!rdr.IsDBNull(pos))
                            bIsCurrent = rdr.GetBoolean(pos);

                        if (bIsCurrent)
                        {
                            pos = rdr.GetOrdinal("DateCompleted");
                            if (!rdr.IsDBNull(pos))
                            {
                                bHasCompleteDate = true;
                                nextDueDate = rdr.GetDateTime(pos);
                                if (nextDueDate.CompareTo(minDueDate) < 0)
                                    minDueDate = nextDueDate;
                                //SNextDueDate = ptnd.NextDueDate.ToString("MM/dd/yyyy");
                            }
                        }

                        pos = rdr.GetOrdinal("Email");
                        if (!rdr.IsDBNull(pos))
                        {
                            email = rdr.GetString(pos);
                        }

                        pos = rdr.GetOrdinal("EmployeeID");
                        if (!rdr.IsDBNull(pos))
                        {
                            empId = rdr.GetString(pos);
                        }

                        pos = rdr.GetOrdinal("NovaStatStrip");
                        if (!rdr.IsDBNull(pos))
                        {
                            bNovaNetCompleted = rdr.GetBoolean(pos);
                        }

                        pos = rdr.GetOrdinal("Vamp");
                        if (!rdr.IsDBNull(pos))
                        {
                            bVampCompleted = rdr.GetBoolean(pos);
                        }
                        
                    }
                    rdr.Close();
                }
                catch (Exception ex)
                {
                    Logger.Error(ex);
                    return null;
                }
            }
            return ptndl;
        }

        static IEnumerable<PostTestNextDue> GetPostTestPeopleFirstDateCompleted(int siteId)
        {
            var ptndl = new List<PostTestNextDue>();

            String strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            using (var conn = new SqlConnection(strConn))
            {
                try
                {
                    var cmd = new SqlCommand("", conn)
                                  {
                                      CommandType = System.Data.CommandType.StoredProcedure,
                                      CommandText = ("GetStaffPostTestsFirstDateCompletedBySite")
                                  };
                    var param = new SqlParameter("@siteID", siteId);
                    cmd.Parameters.Add(param);

                    conn.Open();
                    var rdr = cmd.ExecuteReader();

                    while (rdr.Read())
                    {
                        var ptnd = new PostTestNextDue();

                        var pos = rdr.GetOrdinal("ID");
                        ptnd.Id = rdr.GetInt32(pos);

                        pos = rdr.GetOrdinal("Name");
                        ptnd.Name = rdr.GetString(pos);

                        pos = rdr.GetOrdinal("MinDate");
                        if (!rdr.IsDBNull(pos))
                        {
                            ptnd.NextDueDate = rdr.GetDateTime(pos).AddYears(1);
                            ptnd.SNextDueDate = ptnd.NextDueDate.ToString("MM/dd/yyyy");
                        }

                        pos = rdr.GetOrdinal("Email");
                        if (!rdr.IsDBNull(pos))
                        {
                            ptnd.Email = rdr.GetString(pos);
                        }

                        pos = rdr.GetOrdinal("EmployeeID");
                        if (!rdr.IsDBNull(pos))
                        {
                            ptnd.EmployeeId = rdr.GetString(pos);
                        }

                        pos = rdr.GetOrdinal("NovaStatStrip");
                        if (!rdr.IsDBNull(pos))
                        {
                            ptnd.IsNovaNetTested = rdr.GetBoolean(pos);
                        }

                        pos = rdr.GetOrdinal("Vamp");
                        if (!rdr.IsDBNull(pos))
                        {
                            ptnd.IsVampTested = rdr.GetBoolean(pos);
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
                    Logger.Error(ex);
                    return null;
                }
            }
            return ptndl;
        }
        
        static void WriteNovaNetFile(IEnumerable<NovaNetColumns> lines)
        {
            //write lines to new file
            var fileName = Path.Combine("C:\\Halfpint\\NovaNet\\OperatorsList", "test.csv");
            
            var sw = new StreamWriter(fileName, false);


            sw.WriteLine("NovaNet Operator Import Data,version 2.0,,,,,,,,,");
            foreach (var line in lines)
            {
                sw.Write(line.LastName + ",");
                sw.Write(line.FirstName + ",");
                sw.Write(line.Col3 + ",");
                sw.Write(line.Col4 + ",");
                sw.Write(line.Col5 + ",");
                sw.Write(line.EmployeeId + ",");
                sw.Write(line.Col7 + ",");
                sw.Write(line.Col8 + ",");
                sw.Write(line.Col9 + ",");
                sw.Write(line.StartDate + ",");
                sw.Write(line.EndDate);
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        static List<NovaNetColumns> GetNovaNetFile(string site)
        {
            var di = new DirectoryInfo("C:\\Halfpint\\NovaNet\\OperatorsList");
            var files = di.EnumerateFiles();
            var fileName = "";

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
            var lines = new List<NovaNetColumns>();
            var delimiters = new[] { ',' };
            using (var reader = new StreamReader(fileName))
            {
                int count = 1;
                while (true)
                {
                    var line = reader.ReadLine();
                    if (line == null)
                    {
                        break;
                    }

                    if (count == 1)
                    {                        
                        count = 2;
                        continue;
                    }
                    var cols = new NovaNetColumns();
                    var parts = line.Split(delimiters);

                    cols.LastName = parts[0];
                    cols.FirstName = parts[1];
                    cols.Col3 = parts[2];
                    cols.Col4 = parts[3];
                    cols.Col5 = parts[4];

                    string empId = parts[5];
                    switch(site)
                    {
                        case "CHB":
                            if (empId.Length < 6)
                            {
                                var add = 6 - empId.Length;
                                for (var i = 0; i < add; i++)
                                    empId = "0" + empId;                                    
                            }
                                    
                            break;
                    }
                    cols.EmployeeId = empId;
                    cols.Col7 = parts[6];
                    cols.Col8 = parts[7];
                    cols.Col9 = parts[8];
                    cols.StartDate = parts[9];
                    cols.EndDate = parts[10];

                    lines.Add(cols);
                    // Console.WriteLine("{0} field(s)", parts.Length);
                }
            }
            return lines;
        }
    }

    
    public class SiteInfo
    {
        public int Id { get; set; }
        public string SiteId { get; set; }
        public string Name { get; set; }
        public bool EmpIdRequired { get; set; }
    }

    public class PostTestNextDue
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public DateTime NextDueDate { get; set; }
        public string SNextDueDate { get; set; }
        public string Email { get; set; }
        public string EmployeeId { get; set; }
        public bool IsNovaNetTested { get; set; }
        public bool IsVampTested { get; set; }
        public string Role { get; set; }
    }

    public class NovaNetColumns
    {
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string Col3 { get; set; }
        public string Col4 { get; set; }
        public string Col5 { get; set; }
        public string EmployeeId { get; set; }
        public string Col7 { get; set; }
        public string Col8 { get; set; }
        public string Col9 { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public bool Found{ get ; set ; }
    }

    public class StaffTestsNotCompletedList
    {
        public StaffTestsNotCompletedList()
        {
            TestsNotCompleted = new List<string> {"Overview", "Checks", "NovaStatStrip", "VampJr"};
        }

        public int StaffId { get; set; }
        public string StaffName { get; set; }
        public string Email { get; set; }
        public List<string> TestsNotCompleted { get; set; }
    }
    
    public class SiteEmailLists
    {
        public int SiteId { get; set; }
        public List<PostTestNextDue> DueList { get; set; }
        public List<PostTestNextDue> CompetencyMissingList { get; set; }
        public List<PostTestNextDue> EmailMissingList { get; set; }
        public List<PostTestNextDue> EmployeeIdMissingList { get; set; }
        public List<PostTestNextDue> StaffAddedList { get; set; }
        public List<PostTestNextDue> StaffRemovedList { get; set; }
        public List<StaffTestsNotCompletedList> StaffTestsNotCompletedList { get; set; }
    }
}
