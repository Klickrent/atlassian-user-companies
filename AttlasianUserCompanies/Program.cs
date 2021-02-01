using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace AtlassianUserCompanies
{
    public class Program
    {
        public static string credentials;
        public static Dictionary<string, User> AllUsersDict = new Dictionary<string, User>();
        public static Dictionary<int, Company> AllCompaniesDict = new Dictionary<int, Company>();

        //All Groups which count as licenced for jira and/or confluence
        public static List<AtlassianLicenceGroup> atlassianGroups = new List<AtlassianLicenceGroup> {
            new AtlassianLicenceGroup(AtlassianGrouptypes.Confluence, "1_Administration"),
            new AtlassianLicenceGroup(AtlassianGrouptypes.Confluence, "4_Default"),
            new AtlassianLicenceGroup(AtlassianGrouptypes.Confluence, "administrators"),
            new AtlassianLicenceGroup(AtlassianGrouptypes.Confluence, "atlassian-addons-admin"),
            new AtlassianLicenceGroup(AtlassianGrouptypes.Confluence, "confluence-users"),
            //new AtlassianLicenceGroup(AtlassianGrouptypes.Confluence, "system-administrators"),
            new AtlassianLicenceGroup(AtlassianGrouptypes.Jira, "jira-administrators"),
            new AtlassianLicenceGroup(AtlassianGrouptypes.Jira, "jira-software-users"),
            new AtlassianLicenceGroup(AtlassianGrouptypes.Jira, "jira-users"),
                };

        static void Main(string[] args)
        {
            Console.WriteLine("Atlassian Userstats - Zeppelin");
            Console.WriteLine("");
            Console.WriteLine("Enter Username:");
            string username = Console.ReadLine();
            Console.WriteLine("Enter Password:");
            string passwd = ReadPassword();

            Console.WriteLine("Starting - this will take some minutes.");
            Stopwatch watch = new Stopwatch();
            watch.Start();

            //Convert username und password to the expected jira api format
            credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(username + ":" + passwd));

            // get all Users in Groups which count as licenced for jira and/or confluence
            GetAllLicencedUsers();
            GetAllCompanies();

            // get the company for every licenced user. Too bad we have to query every user one by one.
            // add one to the respective company usercount variables
            Parallel.ForEach(AllUsersDict.Values, (user) =>
            {
                UserCompanyRootObject companyRootObject = JsonResponse<UserCompanyRootObject>($"https://jira.zeppelin.com/rest/stonikbyte-project-team-api/1.0/custom-field-value/{user.key}/1");
                if (companyRootObject.customFieldValueOptions?.Count > 0)
                    user.company = AllCompaniesDict[companyRootObject.customFieldValueOptions[0].CompanyID];
                else
                    //No Company
                    user.company = AllCompaniesDict[-1];

                user.company.UserCount++;
                if (user.jira)
                    user.company.UserCountJira++;
                if (user.confluence)
                    user.company.UserCountConfluence++;
            });

            //Generate the Excel sheet and save it to the desktop
            string filename = @$"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\AtlassianStats{DateTime.Now.Ticks}.xlsx";
            GenerateExcel(filename);

            watch.Stop();
            Console.WriteLine("______________________________________");
            Console.WriteLine();
            Console.WriteLine($"{filename} saved succesfully. Operation took {watch.ElapsedMilliseconds / 1000} seconds.");
            Console.WriteLine("");
            Console.WriteLine("Press Enter to quit");     
            Console.ReadLine();
        }

        public static void GetAllLicencedUsers()
        {
            //Get all active Users (up to 2000)
            var allUsers = JsonResponse<List<User>>
                        ($"https://jira.zeppelin.com/rest/api/latest/user/search?startAt=0&maxResults=1000&username=.");
            AllUsersDict = allUsers.ToDictionary(x => x.key, x => x);
            allUsers = JsonResponse<List<User>>
                        ($"https://jira.zeppelin.com/rest/api/latest/user/search?startAt=1000&maxResults=1000&username=.");
            allUsers.ForEach(x => AllUsersDict.Add(x.key, x));

            Parallel.ForEach(atlassianGroups, (group) =>
            {
                int startAt = 0;
                while (true)
                {
                    JiraGroupRootObject result;
                    try
                    {
                        result = JsonResponse<JiraGroupRootObject>
                        ($"https://jira.zeppelin.com/rest/api/2/group?groupname={group.groupName}&expand=users[{startAt}:{startAt + 49}]");
                    }    
                    catch (Exception)
                    {
                        //Group most likely not present anymore
                        break;
                    }
                    foreach (User user in result.users.items)
                    {
                        if (AllUsersDict.TryGetValue(user.key, out User founduser))
                        {
                            if (group.type == AtlassianGrouptypes.Jira)
                                founduser.jira = true;
                            else
                                founduser.confluence = true;

                            founduser.countLicenceGroups++;
                        }    
                    }
                    if (result.users.endIndex < result.users.size - 1)
                    {
                        startAt = result.users.endIndex++;
                    }
                    else
                    {
                        break;
                    }
                }
            });
        }

        public static void GetAllCompanies()
        {
            CompanyRootObject result = JsonResponse<CompanyRootObject>("https://jira.zeppelin.com/rest/stonikbyte-project-team-api/1.0/custom-field/1");
            foreach (var item in result.companies)
            {
                item.name = WebUtility.HtmlDecode(item.name);
                AllCompaniesDict.Add(item.id, item);
            }
            AllCompaniesDict.Add(-1, new Company { id = -1, name = "--- NO COMPANY ---" });
        }


        public static T JsonResponse<T>(string requestString) where T : new()
        {
            WebRequest request = WebRequest.Create(requestString);
            request.Method = "GET";
            request.ContentType = "application/json; charset=utf-8";
            request.Headers.Add(HttpRequestHeader.Authorization, $"Basic {credentials}");

            var content = string.Empty;

            try
            {
                using (var response = (HttpWebResponse)request.GetResponse())
                {
                    using (var stream = response.GetResponseStream())
                    {
                        using (var sr = new StreamReader(stream))
                        {
                            content = sr.ReadToEnd();
                        }
                    }
                }
            }
            catch (WebException e)
            {
                Console.WriteLine();
                Console.WriteLine(e.Message);

                //detailed error when its an 404
                if ((e.Response as HttpWebResponse)?.StatusCode == HttpStatusCode.NotFound == true)
                {
                    using (Stream stream = e.Response.GetResponseStream())
                    {
                        StreamReader reader = new StreamReader(stream, Encoding.UTF8);
                        Console.WriteLine(reader.ReadToEnd());
                    }
                }

                Console.WriteLine();
                Console.WriteLine("Press Enter to quit.");
                Console.ReadLine();
                Environment.Exit(0);
            }
            
            return JsonSerializer.Deserialize<T>(content);
        }

        //Read password from Console Method -> places * instead of entered characters
        public static string ReadPassword()
        {
            string password = "";
            ConsoleKeyInfo info = Console.ReadKey(true);
            while (info.Key != ConsoleKey.Enter)
            {
                if (info.Key != ConsoleKey.Backspace)
                {
                    Console.Write("*");
                    password += info.KeyChar;
                }
                else if (info.Key == ConsoleKey.Backspace)
                {
                    if (!string.IsNullOrEmpty(password))
                    {
                        // remove one character from the list of password characters
                        password = password.Substring(0, password.Length - 1);
                        // get the location of the cursor
                        int pos = Console.CursorLeft;
                        // move the cursor to the left by one character
                        Console.SetCursorPosition(pos - 1, Console.CursorTop);
                        // replace it with space
                        Console.Write(" ");
                        // move the cursor to the left by one character again
                        Console.SetCursorPosition(pos - 1, Console.CursorTop);
                    }
                }
                info = Console.ReadKey(true);
            }
            // add a new line because user pressed enter at the end of their password
            Console.WriteLine();
            return password;
        }

        //Generate the excel file with two sheets, first for the existing companies and their usercount, second for every single licenced user
        public static void GenerateExcel(string fileName)
        {
            using (var workbook = new XLWorkbook())
            {
                //Sheet 1 => Companies Summary
                var worksheet = workbook.Worksheets.Add("Companies");

                worksheet.Cell("A1").Value = "Company";
                worksheet.Cell("B1").Value = "Users";
                worksheet.Cell("C1").Value = "Jira Users";
                worksheet.Cell("D1").Value = "Confluence Users";

                int i = 0;
                foreach (var item in AllCompaniesDict.Values)
                {
                    worksheet.Cell(i + 2, 1).Value = item.name;
                    worksheet.Cell(i + 2, 2).Value = item.UserCount;
                    worksheet.Cell(i + 2, 3).Value = item.UserCountJira;
                    worksheet.Cell(i + 2, 4).Value = item.UserCountConfluence;
                    i++;
                }

                for (int i2 = 1; i2 < 5; i2++)
                {
                    worksheet.Cell(1, i2).Style.Font.Bold = true;
                    worksheet.Column(i2).AdjustToContents();
                }

                //Sheet 2 => Individual Users
                worksheet = workbook.Worksheets.Add("User");
                worksheet.Cell("A1").Value = "Anzeigename";
                worksheet.Cell("B1").Value = "Name";
                worksheet.Cell("C1").Value = "Email";
                worksheet.Cell("D1").Value = "Compay";
                worksheet.Cell("E1").Value = "Jira";
                worksheet.Cell("F1").Value = "Confluence";

                i = 0;
                foreach (var item in AllUsersDict.Values)
                { 
                    //only insert users with jira and/or confluence license
                    if ( item.countLicenceGroups > 0)
                    {
                        worksheet.Cell(i + 2, 1).Value = item.displayName;
                        worksheet.Cell(i + 2, 2).Value = item.name;
                        worksheet.Cell(i + 2, 3).Value = item.emailAddress;
                        worksheet.Cell(i + 2, 4).Value = item.company.name;
                        if (item.jira)
                            worksheet.Cell(i + 2, 5).Value = "X";
                        if (item.confluence)
                            worksheet.Cell(i + 2, 6).Value = "X";
                        i++;
                    }
                }

                for (int i2 = 1; i2 < 7; i2++)
                {
                    worksheet.Cell(1, i2).Style.Font.Bold = true;
                    worksheet.Column(i2).AdjustToContents();
                }

                //Save Excelfile
                workbook.SaveAs(fileName);
            }
        }
    }
}
