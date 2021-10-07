using System;
using System.Reflection;
using Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Text.RegularExpressions;
/*
* References: 
* https://stackoverflow.com/questions/53737653/how-to-get-data-from-outlook-calendar
* https://stackoverflow.com/questions/90899/net-get-all-outlook-calendar-items
*/

namespace SharpCalendar
{
    class Program
    {
        static void Main(string[] args)
        {
            if ((args.Length != 2) || (Regex.IsMatch(args[0], "^[a-zA-Z]*$")) || (Regex.IsMatch(args[1], "^[a-zA-Z]*$")) || args[0] == "-h")
                {
                    Console.WriteLine(@"
 _____ _                      _____       _                _            
/  ___| |                    /  __ \     | |              | |           
\ `--.| |__   __ _ _ __ _ __ | /  \/ __ _| | ___ _ __   __| | __ _ _ __ 
 `--. \ '_ \ / _` | '__| '_ \| |    / _` | |/ _ \ '_ \ / _` |/ _` | '__|
/\__/ / | | | (_| | |  | |_) | \__/\ (_| | |  __/ | | | (_| | (_| | |   
\____/|_| |_|\__,_|_|  | .__/ \____/\__,_|_|\___|_| |_|\__,_|\__,_|_|   
                       | |                                              
                       |_|                                              
" +
"\n" +
"Developed By: @sadpanda_sec \n\n" +
"Description: Read Outlook Calendar Entries Monthly.\n\n" +
"Examples/Usage:\n" +
"Get Calendar Entries Today + 5 Months ahead:           SharpCalendar.exe 0 5 \n" +
"Get Calendar Entries 1 Month Prior + 3 Months Ahead:   SharpCalendar.exe -1 3");
                    System.Environment.Exit(0);
                }
            else 
                {
                    try
                        {
                            Application oApp = null;
                            NameSpace mapiNamespace = null;
                            MAPIFolder CalendarFolder = null;
                            Items outlookCalendarItems = null;

                            oApp = new Application();
                            mapiNamespace = oApp.GetNamespace("MAPI");
                            CalendarFolder = mapiNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                            outlookCalendarItems = CalendarFolder.Items;
                            outlookCalendarItems.IncludeRecurrences = true;
                            mapiNamespace.Logon(Missing.Value, Missing.Value, false, false);

                            int timeframestart = Convert.ToInt32(args[0]);
                            int timeframeend = Convert.ToInt32(args[1]);

                            DateTime start = DateTime.Today.AddMonths(timeframestart);
                            DateTime end = start.AddMonths(timeframeend);
                            Items range = GetAppointmentsInRange(CalendarFolder, start, end);

                            if (range != null)
                                {
                                    DateTime now = DateTime.Now;
                                    Console.WriteLine("\n" + "Executing SharpCalendar: " + now + " " + TimeZoneInfo.Local);
                                    foreach (AppointmentItem apt in range)
                                        {
                                            string status = apt.BusyStatus.ToString();
                                            if (apt.Body != null)
                                                {
                                                    Console.WriteLine("\n--------------------------------------");
                                                    Console.WriteLine("| Subject: " + apt.Subject);
                                                    Console.WriteLine("| Organizer: " + apt.Organizer);
                                                    Console.WriteLine("| Start: " + apt.Start.ToString());
                                                    Console.WriteLine("| End: " + apt.End.ToString());
                                                    Console.WriteLine("| TimeZone: " + oApp.TimeZones.CurrentTimeZone.Name);
                                                    Console.WriteLine("| Location: " + apt.Location);
                                                    Console.WriteLine("| Recurring: " + apt.IsRecurring);
                                                    Console.WriteLine("| All Day Event: " + apt.AllDayEvent);
                                                    Console.WriteLine("| Office Status: " + status.Substring(2));
                                                    Console.WriteLine("--------------------------------------");
                                                    Console.WriteLine("--------------------------------------");
                                                    Console.WriteLine("|    Body Content of Calendar Entry  |");
                                                    Console.WriteLine("--------------------------------------\n");
                                                    Console.WriteLine(apt.Body + "\n");
                                                }
                                            else
                                                {
                                                    Console.WriteLine("\n--------------------------------------");
                                                    Console.WriteLine("| Subject: " + apt.Subject);
                                                    Console.WriteLine("| Organizer: " + apt.Organizer);
                                                    Console.WriteLine("| Start: " + apt.Start.ToString());
                                                    Console.WriteLine("| End: " + apt.End.ToString());
                                                    Console.WriteLine("| TimeZone: " + oApp.TimeZones.CurrentTimeZone.Name);
                                                    Console.WriteLine("| Location: " + apt.Location);
                                                    Console.WriteLine("| Recurring: " + apt.IsRecurring);
                                                    Console.WriteLine("| All Day Event: " + apt.AllDayEvent);
                                                    Console.WriteLine("| Office Status: " + status.Substring(2));
                                                    Console.WriteLine("--------------------------------------");
                                                }
                                        }
                                }
                        mapiNamespace.Logoff();
                        oApp = null;
                        mapiNamespace = null;
                        CalendarFolder = null;
                        outlookCalendarItems = null;
                        System.Environment.Exit(0);

                }
                    catch (System.Exception e)
                    {
                        Console.WriteLine("{0} Exception Caught.", e);
                    }

                }

             Items GetAppointmentsInRange(MAPIFolder calendarFolder, DateTime startTime, DateTime endTime)
                {
                    string filter = "[Start] >= '"
                        + startTime.ToString("g")
                        + "' AND [End] <= '"
                        + endTime.ToString("g") + "'";
                    Debug.WriteLine(filter);
                    try
                    {
                        Items calItems = calendarFolder.Items;
                        calItems.IncludeRecurrences = true;
                        calItems.Sort("[Start]", Type.Missing);
                        Items restrictItems = calItems.Restrict(filter);
                        if (restrictItems.Count > 0)
                        {
                            return restrictItems;
                        }
                        else
                        {
                            return null;
                        }
                    }
                    catch
                    {
                        return null;
                    }
                }
            }
        }
    }