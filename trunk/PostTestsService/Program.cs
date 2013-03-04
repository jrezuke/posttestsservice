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
        private static bool _bSendEmails = true;

        static void Main(string[] args)
        {       
            Logger.Info("Starting PostTests Service");
            
            if (args.Length > 0)
            {
                _bSendEmails = false;
                Console.WriteLine("Argument:" + args[0]);
                Logger.Info("Argument:" + args[0]);
            }
            
            var path = AppDomain.CurrentDomain.BaseDirectory;
            
            //get sites 
            var sites = GetSites();
            
            //iterate sites
            foreach (var si in sites)
            {
                if (!si.EmpIdRequired)
                    continue;

                si.SiteEmailLists = new SiteEmailLists
                                        {
                                            SiteId = si.Id,
                                            DueList = new List<PostTestNextDue>(),
                                            CompetencyMissingList = new List<PostTestNextDue>(),
                                            EmailMissingList = new List<PostTestNextDue>(),
                                            EmployeeIdMissingList = new List<PostTestNextDue>()
                                        };

                
                Console.WriteLine(si.Name);
                Logger.Info("For Site:" + si.Name + " - " + si.SiteId);
                Logger.Debug("For Site:" + si.Name + " - " + si.SiteId);
                
                //Get staff info including next due date - next due date will be 1 year from today for new staff
                //staff roles not included are Admin, DCC , Nurse generic (nurse accounts with a user name)
                si.PostTestNextDues = GetAllStaffPostTestsCompleted(si.Id); //GetPostTestPeopleFirstDateCompleted(si.Id);
                
                //this is used to build the tests not completed lists - only includes staff that can be added to the list
                si.PostTestNextDues2 = new List<PostTestNextDue>();

                //iterate people                
                foreach (var postTestNextDue in si.PostTestNextDues)
                {
                    var bContinue = false;
                    Console.WriteLine(postTestNextDue.Name + ":" + postTestNextDue.SNextDueDate + ", email: " + postTestNextDue.Email + ", Employee ID: " + postTestNextDue.EmployeeId + ", Role: " + postTestNextDue.Role);
                    
                    
                    
                    //make sure they are nova net certified
                    if ((!postTestNextDue.IsNovaNetTested) || (!postTestNextDue.IsVampTested))
                    {
                        Logger.Info("Competency needed for " + postTestNextDue.Name);
                        si.SiteEmailLists.CompetencyMissingList.Add(postTestNextDue);
                        bContinue = true;
                    }

                    if (postTestNextDue.Email == null)
                    {
                        Logger.Info("Email missing for " + postTestNextDue.Name);
                        si.SiteEmailLists.EmailMissingList.Add(postTestNextDue);
                        bContinue = true;
                    }
                    else 
                    {
                        if (postTestNextDue.Email.Trim().Length == 0)
                        {
                            Logger.Info("Email missing for " + postTestNextDue.Name);
                            si.SiteEmailLists.EmailMissingList.Add(postTestNextDue);
                            bContinue = true;
                        }
                    }
                    if (si.EmpIdRequired)
                    {
                        if (postTestNextDue.EmployeeId == null)
                        {
                            Logger.Info("Employee ID missing for " + postTestNextDue.Name);
                            si.SiteEmailLists.EmployeeIdMissingList.Add(postTestNextDue);
                            bContinue = true;
                        }
                        else
                        {
                            if (postTestNextDue.EmployeeId.Trim().Length == 0)
                            {
                                Logger.Info("Employee ID missing for " + postTestNextDue.Name);
                                si.SiteEmailLists.EmployeeIdMissingList.Add(postTestNextDue);
                                bContinue = true;
                            }
                        }
                    }
                    if (bContinue)
                        continue;

                    //see if all required post tests are completed
                    

                    //add to list2 for 2nd run
                    si.PostTestNextDues2.Add(postTestNextDue);
                    
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
                    
                    //todo remove this for production
                    //int retVal = SetPostTestsCompletedIsCurrent(postTestNextDue.Id);
                    //Logger.Info("Number of tests set IsCurrent=0: " + retVal);
                            
                    //send email to user                                                           
                    var to = new[] { postTestNextDue.Email };

                    const string subject = "Annual Halfpint Post Tests Due";
                    const string body = "Your annual halpint post tests are now available at the link below.  Please complete the required tests as soon as possible.";

                    if(_bSendEmails)
                        SendHtmlEmail(subject, to, null, body, path, @"<a href='http://halfpintstudy.org/hpProd/PostTests/Initialize'>Halfpint Study Post Tests</a>");
                            
                    //add to list - to be sent to coordinator
                    si.SiteEmailLists.DueList.Add(postTestNextDue);
                } //foreach (var ptnd in ptndl)
                
            } //foreach (var si in sites)
            
            //now update the nova net files
            Console.WriteLine("-------------------------");
            Console.WriteLine("Updating nova net files");
            Console.WriteLine("-------------------------");
            
            //iterate sites
            foreach (var si in sites)
            {
                //skip for sites not needed
                if (!si.EmpIdRequired)
                    continue;

                //var siteEmailList = siteLists.Find(x => x.SiteId == si.Id);

                si.SiteEmailLists.StaffAddedList = new List<PostTestNextDue>();
                si.SiteEmailLists.StaffRemovedList = new List<PostTestNextDue>();
                
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
                Logger.Info("sitePtnd2List.Find(x => x.SiteId == si.Id)");
                foreach (var ptnd in si.PostTestNextDues2)
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
                        //var lineDate = DateTime.Parse(line.EndDate);
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
                        si.SiteEmailLists.StaffAddedList.Add(ptnd);
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
                WriteNovaNetFile(lines, si.Name);
                Logger.Info("WriteNovaNetFile:" + si.Name);
            }//foreach (var si in sites) - write file

            if (_bSendEmails)
            {
                foreach (var si in sites)
                {
                    //skip for sites not needed
                    if (!si.EmpIdRequired)
                        continue;

                    //var siteEmailList = siteLists.Find(x => x.SiteId == si.Id);
                    si.SiteEmailLists.StaffTestsNotCompletedList = GetStaffPostTestCompleted(si.Id).ToList();
                    SendCoordinatorsEmail(si.Id, si.Name, si.SiteEmailLists, path);
                    
                } //foreach (var si in sites) - tests not completed
            }
            Console.Read();
        }


        internal static void SendCoordinatorsEmail(int site, string siteName, SiteEmailLists siteEmailLists, string path)
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
            
            SendHtmlEmail("Post Tests Notifications - " + siteName, coordinators.Select(coord => coord.Email).ToArray(), null, sbBody.ToString(), path, @"<a href='http://halfpintstudy.org/hpProd/'>Halfpint Study Website</a>");
        }

        internal static void SendHtmlEmail(string subject, string[] toAddress, string[] ccAddress, string bodyContent, string appPath, string url, string bodyHeader = "")
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
            Console.WriteLine("Subject:" + subject);
            Console.Write("To:" + toAddress[0]);
            Console.Write("Email:" + sb);

            var smtp = new SmtpClient();
            smtp.Send(mm);
        }

        internal static List<MembershipUser> GetUserInRole(string role, int site)
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

        public static List<String> GetActiveTests()
        {
            var list = new List<string>();
            var strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();

            using (var conn = new SqlConnection(strConn))
            {
                try
                {
                    var cmd = new SqlCommand("", conn)
                                  {
                                      CommandType = System.Data.CommandType.Text,
                                      CommandText = "SELECT Name FROM PostTests WHERE Active=1"
                                  };

                    conn.Open();
                    var rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        list.Add(rdr.GetString(0));
                    }
                    rdr.Close();
                }
                catch (Exception ex)
                {
                    Logger.Error(ex);
                }
            }

            return list;
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
            //bool anyTestsNotCompleted = false;
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
                        int pos = rdr.GetOrdinal("ID");
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
                    //capture the last one
                    if (staffId > 0)
                    {
                        if (tncl.TestsNotCompleted.Count > 0)
                            tncll.Add(tncl);
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

        static List<PostTestNextDue> GetAllStaffPostTestsCompleted(int siteId)
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
                    //string sMinDate = null;
                    string name = null;
                    string email = null;
                    string empId = null;

                    var minDueDate = DateTime.Parse("01/01/2100");
                    var bHasCompleteDate = false;
                    var bVampCompleted = false;
                    var bNovaNetCompleted = false;
                    var role = String.Empty;
                    var bIsNew = false;

                    var rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var pos = rdr.GetOrdinal("Role");
                        role = rdr.GetString(pos);
                        var bIsCurrent = false;
                        
                        if (role == "Admin" )
                            continue;

                        var userName = string.Empty;
                        pos = rdr.GetOrdinal("UserName");
                        if(!rdr.IsDBNull(pos))
                            userName = rdr.GetString(pos);

                        //skip generic roles
                        if (role == "Nurse" || role == "DCC")
                        {
                            if(userName !=  string.Empty)
                                continue;
                        }
                        
                        pos = rdr.GetOrdinal("ID");
                        var iD = rdr.GetInt32(pos);
                        
                        if (iD != curId)
                        {
                            //add to list if current id changes
                            if (curId != 0)
                            {
                                var ptnd = new PostTestNextDue
                                               {
                                                   Id = curId,
                                                   Name = name,
                                                   Role = role,
                                                   Email = email,
                                                   EmployeeId = empId,
                                                   IsNovaNetTested = bNovaNetCompleted,
                                                   IsVampTested = bVampCompleted
                                               };
                                if (bHasCompleteDate)
                                {
                                    ptnd.NextDueDate = minDueDate.AddYears(1);
                                    ptnd.SNextDueDate = ptnd.NextDueDate.ToString("MM/dd/yyyy");
                                }
                                ptnd.IsNew = bIsNew;

                                ptndl.Add(ptnd);
                            }
                            curId = iD;
                            minDueDate = DateTime.Parse("01/01/2100");
                            bHasCompleteDate = false;
                            bVampCompleted = false;
                            bNovaNetCompleted = false;
                            bIsNew = false;
                            email = null;
                            empId = null;
                        }
                        
                        pos = rdr.GetOrdinal("Name");
                        name = rdr.GetString(pos);

                        pos = rdr.GetOrdinal("IsCurrent");
                        if (rdr.IsDBNull(pos))
                            bIsNew = true;
                        else
                        {
                            bIsNew = false;
                            bIsCurrent = rdr.GetBoolean(pos);
                        }

                        if (bIsCurrent)
                        {
                            pos = rdr.GetOrdinal("DateCompleted");
                            if (!rdr.IsDBNull(pos))
                            {
                                bHasCompleteDate = true;
                                var nextDueDate = rdr.GetDateTime(pos);
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
                    
                    //add the last staff
                    if (curId != 0)
                    {
                        var ptnd = new PostTestNextDue
                                       {
                                           Id = curId,
                                           Name = name,
                                           Role = role,
                                           Email = email,
                                           EmployeeId = empId,
                                           IsNovaNetTested = bNovaNetCompleted,
                                           IsVampTested = bVampCompleted,
                                           IsNew = bIsNew
                                       };
                        if (bHasCompleteDate)
                        {
                            ptnd.NextDueDate = minDueDate.AddYears(1);
                            ptnd.SNextDueDate = ptnd.NextDueDate.ToString("MM/dd/yyyy");
                        }
                        ptndl.Add(ptnd);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error(ex);
                    return null;
                }
            }
            return ptndl;
        }

        //static IEnumerable<PostTestNextDue> GetPostTestPeopleFirstDateCompleted(int siteId)
        //{
        //    var ptndl = new List<PostTestNextDue>();

        //    String strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
        //    using (var conn = new SqlConnection(strConn))
        //    {
        //        try
        //        {
        //            var cmd = new SqlCommand("", conn)
        //                          {
        //                              CommandType = System.Data.CommandType.StoredProcedure,
        //                              CommandText = ("GetStaffPostTestsFirstDateCompletedBySite")
        //                          };
        //            var param = new SqlParameter("@siteID", siteId);
        //            cmd.Parameters.Add(param);

        //            conn.Open();
        //            var rdr = cmd.ExecuteReader();

        //            while (rdr.Read())
        //            {
        //                var ptnd = new PostTestNextDue();

        //                var pos = rdr.GetOrdinal("ID");
        //                ptnd.Id = rdr.GetInt32(pos);

        //                pos = rdr.GetOrdinal("Name");
        //                ptnd.Name = rdr.GetString(pos);

        //                pos = rdr.GetOrdinal("MinDate");
        //                if (!rdr.IsDBNull(pos))
        //                {
        //                    ptnd.NextDueDate = rdr.GetDateTime(pos).AddYears(1);
        //                    ptnd.SNextDueDate = ptnd.NextDueDate.ToString("MM/dd/yyyy");
        //                }

        //                pos = rdr.GetOrdinal("Email");
        //                if (!rdr.IsDBNull(pos))
        //                {
        //                    ptnd.Email = rdr.GetString(pos);
        //                }

        //                pos = rdr.GetOrdinal("EmployeeID");
        //                if (!rdr.IsDBNull(pos))
        //                {
        //                    ptnd.EmployeeId = rdr.GetString(pos);
        //                }

        //                pos = rdr.GetOrdinal("NovaStatStrip");
        //                if (!rdr.IsDBNull(pos))
        //                {
        //                    ptnd.IsNovaNetTested = rdr.GetBoolean(pos);
        //                }

        //                pos = rdr.GetOrdinal("Vamp");
        //                if (!rdr.IsDBNull(pos))
        //                {
        //                    ptnd.IsVampTested = rdr.GetBoolean(pos);
        //                }

        //                pos = rdr.GetOrdinal("Role");
        //                if (!rdr.IsDBNull(pos))
        //                {
        //                    ptnd.Role = rdr.GetString(pos);
        //                }
        //                ptndl.Add(ptnd);
        //            }
        //            rdr.Close();
        //        }
        //        catch (Exception ex)
        //        {
        //            Logger.Error(ex);
        //            return null;
        //        }
        //    }
        //    return ptndl;
        //}
        
        static void WriteNovaNetFile(IEnumerable<NovaNetColumns> lines, string siteName)
        {
            //write lines to new file
            var folderPath = ConfigurationManager.AppSettings["StatStripListPath"];
            var fileName = siteName + " " + ConfigurationManager.AppSettings["StatStripListName"];

            var fullpath = Path.Combine(folderPath, fileName);


            var sw = new StreamWriter(fullpath, false);


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
            var folderPath = ConfigurationManager.AppSettings["StatStripListPath"];
            var di = new DirectoryInfo(folderPath);
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

            fileName = Path.Combine(folderPath, fileName);
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

        //public static List<PostTest> GetStaffPostTestsCompletedCurrentAndActive(string staffId)
        //{
        //    var tests = new List<PostTest>();

        //    String strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
        //    using (var conn = new SqlConnection(strConn))
        //    {
        //        try
        //        {
        //            var cmd = new SqlCommand("", conn)
        //                          {
        //                              CommandType = System.Data.CommandType.StoredProcedure,
        //                              CommandText = ("GetPostTestsActive")
        //                          };

        //            conn.Open();
        //            var rdr = cmd.ExecuteReader();
        //            while (rdr.Read())
        //            {
        //                var test = new PostTest();

        //                var pos = rdr.GetOrdinal("ID");
        //                test.Id = rdr.GetInt32(pos);

        //                pos = rdr.GetOrdinal("Name");
        //                test.Name = rdr.GetString(pos);

        //                pos = rdr.GetOrdinal("PathName");
        //                test.PathName = rdr.GetString(pos);

        //                test.sDateCompleted = "";

        //                tests.Add(test);
        //            }
        //            rdr.Close();
        //            conn.Close();

        //            cmd = new SqlCommand("", conn)
        //                      {
        //                          CommandType = System.Data.CommandType.StoredProcedure,
        //                          CommandText = ("GetStaffPostTestsCompletedCurrentAndActive")
        //                      };
        //            var param = new SqlParameter("@staffId", staffId);
        //            cmd.Parameters.Add(param);

        //            conn.Open();
        //            rdr = cmd.ExecuteReader();
        //            while (rdr.Read())
        //            {
        //                var pos = rdr.GetOrdinal("TestID");
        //                var testId = rdr.GetInt32(pos);
        //                var test = tests.Find(x => x.Id == testId);

        //                pos = rdr.GetOrdinal("DateCompleted");
        //                test.DateCompleted = rdr.GetDateTime(pos);
        //                test.sDateCompleted = (test.DateCompleted != null
        //                                           ? test.DateCompleted.Value.ToString("MM/dd/yyyy")
        //                                           : "");
        //                test.IsCompleted = true;

        //            }
        //            rdr.Close();
        //        }
        //        catch (Exception ex)
        //        {
        //            Logger.Error(ex);
        //            return null;
        //        }
        //    }
        //    return tests;
        //}
    }

    //public class PostTest
    //{
    //    public int Id { get; set; }
    //    public string Name { get; set; }
    //    public string PathName { get; set; }
    //    public DateTime? DateCompleted { get; set; }
    //    public string sDateCompleted { get; set; }
    //    public bool IsCompleted { get; set; }
    //}
    
    public class SiteInfo
    {
        public int Id { get; set; }
        public string SiteId { get; set; }
        public string Name { get; set; }
        public bool EmpIdRequired { get; set; }
        public SiteEmailLists SiteEmailLists { get; set; }
        public List<PostTestNextDue> PostTestNextDues { get; set; }
        public List<PostTestNextDue> PostTestNextDues2 { get; set; } 
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
        public bool IsNew { get; set; }
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
            TestsNotCompleted = Program.GetActiveTests();
        }

        public int StaffId { get; set; }
        public string StaffName { get; set; }
        public string Email { get; set; }
        public List<string> TestsNotCompleted { get; set; }
    }
    
    public class SitePostTestDueList
    {
        public int SiteId { get; set; }
        public List<PostTestNextDue> PostTestNextDueList { get; set; }
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
