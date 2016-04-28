using System;
using System.Collections.Specialized;
using System.Linq;
using System.Windows.Forms;
using System.Text;
using System.Text.RegularExpressions;

using CBLabs.CybozuConnect;

namespace Cybozu.CrossSync
{
    static class Program
    {
        public const string DescriptionHeaderName = "# CrossSync"; // Don't modify this.

        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new SettingForm());
        }

        public static bool IsConfigured(Properties.Settings settings)
        {
            if (string.IsNullOrEmpty(settings.FirstUrl) || string.IsNullOrEmpty(settings.SecondUrl)) return false;
            if (string.IsNullOrEmpty(settings.FirstPostfix) || string.IsNullOrEmpty(settings.SecondPostfix)) return false;
            
            return true;
        }

        public static bool CanSync(out CybozuException ex)
        {
            ex = null;

            Properties.Settings settings = Properties.Settings.Default;
            if (!IsConfigured(settings)) return false;

            App firstApp, secondApp;
            Schedule firstSchedule, secondSchedule;

            try
            {
                firstApp = new App(settings.FirstUrl);
                firstApp.Auth(settings.FirstUsername, settings.FirstPassword);
                firstSchedule = new Schedule(firstApp);

                secondApp = new App(settings.SecondUrl);
                secondApp.Auth(settings.SecondUsername, settings.SecondPassword);
                secondSchedule = new Schedule(secondApp);
            }
            catch (CybozuException e)
            {
                // fail to auth
                ex = e;
                return false;
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        public static bool Sync()
        {
            Properties.Settings settings = Properties.Settings.Default;
            if (!IsConfigured(settings)) return false;

            App firstApp = new App(settings.FirstUrl);
            firstApp.Auth(settings.FirstUsername, settings.FirstPassword);
            Schedule firstSchedule = new Schedule(firstApp);

            App secondApp = new App(settings.SecondUrl);
            secondApp.Auth(settings.SecondUsername, settings.SecondPassword);
            Schedule secondSchedule = new Schedule(secondApp);

            // sync span
            DateTime start = DateTime.Now.Date;
            DateTime end = start.AddMonths(1);

            // current events in first
            ScheduleEventCollection event1to2 = new ScheduleEventCollection();
            ScheduleEventCollection event1from2 = new ScheduleEventCollection();
            GetEvents(firstApp, firstSchedule, settings.SecondPostfix, start, end, event1to2, event1from2);

            // current events in second
            ScheduleEventCollection event2to1 = new ScheduleEventCollection();
            ScheduleEventCollection event2from1 = new ScheduleEventCollection();
            GetEvents(secondApp, secondSchedule, settings.FirstPostfix, start, end, event2to1, event2from1);

            // update modified events
            UpadteModifiedEvents(secondSchedule, firstSchedule, event1to2, event2from1, settings.FirstPostfix);
            UpadteModifiedEvents(firstSchedule, secondSchedule, event2to1, event1from2, settings.SecondPostfix);

            // remove old copied events
            RemoveInvalidCopiedEvents(secondSchedule, event2from1);
            RemoveInvalidCopiedEvents(firstSchedule, event1from2);

            // add new copied events
            CopyValidEvents(secondSchedule, firstSchedule, event1to2, settings.FirstPostfix);
            CopyValidEvents(firstSchedule, secondSchedule, event2to1, settings.SecondPostfix);

            settings.LastSynchronized = DateTime.Now.ToString("o");
            settings.Save();

            return true;
        }

        public static void GetEvents(App app, Schedule schedule, string postfix, DateTime start, DateTime end, ScheduleEventCollection eventTo, ScheduleEventCollection eventFrom)
        {
            DateTime marginStart = start.AddDays(-1.0);
            DateTime marginEnd = end.AddDays(1.0);
            ScheduleEventCollection eventList = schedule.GetEventsByTarget(marginStart, marginEnd, Schedule.TargetType.User, app.UserId);

            Properties.Settings settings = Properties.Settings.Default;

            foreach (ScheduleEvent scheduleEvent in eventList)
            {
                if (scheduleEvent.StartOnly)
                {
                    if (scheduleEvent.Start.CompareTo(start) < 0) continue;
                }
                else if (scheduleEvent.AllDay || scheduleEvent.IsBanner)
                {
                    if (scheduleEvent.End.CompareTo(start) < 0) continue;
                }
                else
                {
                    if (scheduleEvent.Start.Equals(scheduleEvent.End))
                    {
                        if (scheduleEvent.End.CompareTo(start) < 0) continue;
                    }
                    else
                    {
                        if (scheduleEvent.End.CompareTo(start) <= 0) continue;
                    }
                }
                if (scheduleEvent.Start.CompareTo(end) >= 0) continue;

                if (scheduleEvent.Detail.EndsWith(postfix) || scheduleEvent.Description.StartsWith(DescriptionHeaderName))
                {
                    //if (scheduleEvent.MemberCount <= 1)
                    {
                        eventFrom.Add(scheduleEvent);
                    }
                }
                else if (!scheduleEvent.AllDay || (settings.AllDay && scheduleEvent.AllDay) || (settings.Banner && scheduleEvent.IsBanner) || (settings.Temporary && scheduleEvent.IsTemporary) || (settings.Private && scheduleEvent.IsPrivate) || (settings.Qualified && scheduleEvent.IsQualified))
                {
                    eventTo.Add(scheduleEvent);
                }
            }
        }

        public static void UpadteModifiedEvents(Schedule schedule, Schedule scheduleSrc, ScheduleEventCollection srcEventList, ScheduleEventCollection destEventList, string postfix)
        {
            try
            {
                UpadteModifiedEventsMain(schedule, scheduleSrc, srcEventList, destEventList, postfix, true);
            }
            catch (CybozuException e)
            {
                if (e.Code == "14312" || e.Code == "GRN_SCHD_13208")
                {
                    UpadteModifiedEventsMain(schedule, scheduleSrc, srcEventList, destEventList, postfix, false);
                }
                else
                {
                    throw;
                }
            }
        }

        public static void UpadteModifiedEventsMain(Schedule schedule, Schedule scheduleSrc, ScheduleEventCollection srcEventList, ScheduleEventCollection destEventList, string postfix, bool facilitySync)
        {
            Regex reg = new Regex(string.Format(@"^{0}\((?<id>\d*),(?<version>\d*)\): ", DescriptionHeaderName));

            ScheduleEventCollection modifiedEventsList = new ScheduleEventCollection();
            for (int i = destEventList.Count - 1; i >= 0; i--)
            {
                ScheduleEvent destEvent = destEventList[i];
                Match match = reg.Match(destEvent.Description);

                if (!match.Success) continue;
                string savedId = match.Groups["id"].Value;
                string savedVersion = match.Groups["version"].Value;

                ScheduleEvent srcEvent = srcEventList.FirstOrDefault<ScheduleEvent>(elem => elem.ID == savedId && elem.Start.Date.Equals(destEvent.Start.Date));
                if (srcEvent == null || srcEvent.Version == savedVersion) continue;

                ScheduleEvent modifiedEvent = CreateCopyEvent(schedule, scheduleSrc, srcEvent, postfix, facilitySync);
                modifiedEvent.ID = destEvent.ID;
                modifiedEvent.Version = destEvent.Version;
                modifiedEventsList.Add(modifiedEvent);

                srcEventList.Remove(srcEvent);
                destEventList.Remove(destEvent);
            }

            if (modifiedEventsList.Count == 0) return;

            schedule.ModifyEvents(modifiedEventsList);
        }
        
        public static void RemoveInvalidCopiedEvents(Schedule schedule, ScheduleEventCollection eventList)
        {
            if (eventList.Count == 0) return;

            StringCollection idList = new StringCollection();
            foreach (ScheduleEvent destEvent in eventList)
            {
                if (!string.IsNullOrEmpty(destEvent.ID))
                {
                    idList.Add(destEvent.ID);
                }
            }

            if (idList.Count == 0) return;

            schedule.RemoveEvents(idList);
        }

        public static void CopyValidEvents(Schedule schedule, Schedule scheduleSrc, ScheduleEventCollection eventList, string postfix)
        {
            try
            {
                CopyValidEventsMain(schedule, scheduleSrc, eventList, postfix, true);
            }
            catch (CybozuException e)
            {
                if (e.Code == "14312" || e.Code == "GRN_SCHD_13208")
                {
                    CopyValidEventsMain(schedule, scheduleSrc, eventList, postfix, false);
                }
                else
                {
                    throw;
                }
            }
        }

        public static void CopyValidEventsMain(Schedule schedule, Schedule scheduleSrc, ScheduleEventCollection eventList, string postfix, bool facilitySync)
        {
            if (eventList.Count == 0) return;

            ScheduleEventCollection newEventsList = new ScheduleEventCollection();
            foreach (ScheduleEvent srcEvent in eventList)
            {
                newEventsList.Add(CreateCopyEvent(schedule, scheduleSrc, srcEvent, postfix, facilitySync));
            }

            if (newEventsList.Count == 0) return;

            schedule.AddEvents(newEventsList);
        }

        public static ScheduleEvent CreateCopyEvent(Schedule schedule, Schedule scheduleSrc, ScheduleEvent srcEvent, string postfix, bool facilitySync)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(DescriptionHeaderName);
            sb.Append(String.Format("({0},{1}): ", srcEvent.ID, srcEvent.Version));
            sb.AppendLine(Resources.CrossSyncDescription);
            sb.AppendLine(scheduleSrc.GetMobileViewURL(srcEvent));
            if (!string.IsNullOrEmpty(srcEvent.Description))
            {
                sb.AppendLine();
                sb.Append(srcEvent.Description);
            }

            ScheduleEvent newEvent = new ScheduleEvent();
            newEvent.EventType = srcEvent.IsBanner ? ScheduleEventType.Banner : ScheduleEventType.Normal;
            newEvent.PublicType = srcEvent.IsPublic ? SchedulePublicType.Public : SchedulePublicType.Private;
            newEvent.Start = srcEvent.Start;
            newEvent.End = srcEvent.End;
            newEvent.AllDay = srcEvent.AllDay;
            newEvent.StartOnly = srcEvent.StartOnly;
            newEvent.Plan = srcEvent.Plan;
            newEvent.Detail = srcEvent.Detail + postfix;
            if (srcEvent.FacilityIds.Count > 0)
            {
                if (facilitySync)
                {
                    foreach (string facilityId in srcEvent.FacilityIds)
                    {
                        Facility facility = schedule.Facilities.SearchFacility(scheduleSrc.Facilities[facilityId].Name);
                        if (facility != null)
                        {
                            newEvent.FacilityIds.Add(facility.Key);
                        }
                    }
                }
                else
                {
                    sb.AppendLine();
                    sb.Append(Resources.FacilityNoSync);
                    string sep = "";
                    foreach (string facilityId in srcEvent.FacilityIds)
                    {
                        sb.Append(sep);
                        sep = ", ";
                        sb.Append(scheduleSrc.Facilities[facilityId].Name);
                    }
                }
            }
            newEvent.Description = sb.ToString();
            newEvent.UserIds.Add(schedule.App.UserId);

            return newEvent;
        }
    }
}
