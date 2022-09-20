using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using NightlyRouteToSlack.Utilities;

namespace NightlyRouteToSlack
{
    class Program
    {
        static void Main(string[] args)
        {
            UtilityClass uc = new UtilityClass();
            JunkUtilities ju = new JunkUtilities();
            MoveUtilities mu = new MoveUtilities();
            ShackUtilities su = new ShackUtilities();
            ConfigSettings cs = new ConfigSettings();

            switch (args[0].ToString())
            {
                //1-800-GOT-JUNK? ******************
                case "gj_import_employee_hours_worked":
                    ju.RunGJImportEmployeeHoursWorked();
                    break;

                case "gj_daily_employee_jobs_done":
                    ju.RunGJDailyEmployeeJobsDone();
                    break;

                case "gj_daily_hours_worked_dump":
                    ju.RunGJDailyHoursWorkedDump();
                    break;  

                case "gj_nps_scorecard":
                    ju.RunGJNPSScorecard();
                    break;

                case "junk_sales_lead_ajs":
                    ju.RunSalesLeadAJS();
                    break;

                case "gj_daily_scorecard":
                    ju.JunkScorecardDaily();
                    break;

                case "daily_recon":
                    ju.JunkDailyRecon();
                    break;

                case "daily_recon_west":
                    ju.JunkDailyRecon();
                    break;

                case "daily_schedule":                              //#gotjundailyschedule
                    ju.RunDailySchedule(false);
                    break;

                case "daily_schedule_test":                         //#gotjundailyschedule
                    ju.RunDailySchedule(true);
                    break;

                case "junk_visa_transactions":                      //#gotjunkvisa
                    ju.JunkReportOutVisaTransactions();
                    break;

                case "junk_weekly_scorecard":                       //#gotjunkscorecard
                    ju.JunkReportOutWeeklyScorecard();
                    break;

                case "shack_update_ot_awareness":                   //run every Sunday at 5am
                    su.ShackShineUpdateOTAwareness();
                    break;

                case "junk_ot_awareness":                           //run every day at 5:30am
                    ju.JunkOTAwarenessUpdate();
                    break;

                case "junk_daily_health_check_import":
                    ju.RunDailyHealthCheckImport();
                    break;

                case "get_square_transactions":
                    ju.ImportSquareTransactions(cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkSquareLocationID"));                 //imports square data into routes database every 10 minutes
                    break;

                case "get_square_transactions_west":
                    ju.ImportSquareTransactions(cs.ReturnConfigSetting("NightlyRouteToSlack", "JunkWestSquareLocationID"));             //imports square data into routes database every 10 minutes
                    break;

                case "download_square_transactions":
                    ju.DownloadSquareTransactions();                //imports square data into junk route journal nightly
                    break;

                case "download_square_transactions_west":
                    ju.DownloadSquareTransactionsWest();            //imports square data into junk west route journal nightly
                    break;

















                //YOU MOVE ME ******************
                case "move_jake_timecard_entry":
                    mu.MoveCheckJakeEnteringHours();
                    break;

                case "move_sales_center_wages":
                    mu.MoveSalesCenterWages();
                    break;

                case "move_sales_center_conversion":
                    mu.MoveSalesCenterConversion();
                    break;

                case "ym_daily_scorecard":
                    mu.MoveDailyScorecard();
                    break;

                case "daily_schedule_move":
                    mu.RunDailySchedule(false);
                    break;

                case "daily_schedule_move_test":
                    mu.RunDailySchedule(true);
                    break;

                case "move_visa_transactions":
                    mu.MoveReportOutVisaTransactions();
                    break;

                case "move_weekly_scorecard":
                    mu.MoveReportOutWeeklyScorecard();
                    break;

                case "move_damage":
                    mu.MoveDamageSlackOutput();
                    break;

                case "download_square_transactions_move":                   //nightly download of square into move route journal
                    mu.DownloadSquareTransactions();
                    break;

                case "move_updated_daily_health_check_sheet":
                    mu.RunDailyHealthCheckImport();
                    break;

                case "update_reviews_received":
                    mu.RunUpdateReviewsReceived();
                    break;

                case "move_bag_drop_roi":
                    mu.RunMoveBagDropROI();
                    break;





                //SHACK SHINE ******************
                case "shack_weekly_scorecard":
                    su.ShackReportOutWeeklyScorecard();
                    break;

                case "ss_daily_scorecard":
                    su.ShackDailyScorecard();
                    break;

                case "shack_visa_transactions":
                    su.ShackReportOutVisaTransactions();
                    break;

                case "ss_nps_tracker":
                    su.ShackNPSTracker();
                    break;

                case "download_square_transactions_shack":
                    su.DownloadSquareTransactions();
                    break;

                case "shack_send_messages":
                    su.ShackSendMessages();
                    break;

                case "ss_daily_checklist":                  
                    su.ShackDailyChecklist();
                    break;



                //360WOW INC *****************************
                case "building_security_next_day":
                    uc.SendToBuildingSecurity();
                    break;

                case "hiring_update":
                    uc.UploadHiringUpdate();
                    break;

                case "covid-tracking":
                    uc.ReportOutCovidTracking();
                    break;

                case "upload_slack_users":
                    uc.UploadSlackUsers();                  
                    break;










                //NOT USED *****************************

                case "daily_wage_shack":                    //not used
                    su.RunDailyWagePercentageShack();
                    break;

                case "daily_route_wages_shack":             //not used
                    su.RunDailyRouteWagesShack();
                    break;

                case "daily_sales_shack":                   //not used
                    su.SendShackSalesPercentageToSlack();
                    break;

                case "square_swipe_percentage_shack":       //not used
                    su.ShackSquareSwipePercentage();
                    break;

                case "shack_check_uniform_inventory":       //not used
                    su.ShackCheckUniformInventory();
                    break;


                case "to_send_messages":
                    //uc.SendToSendMessages();                //not used
                    break;

                case "ym_daily_checklist":                      //not used
                    mu.MoveDailyChecklist();
                    break;

                case "daily_wage_move":                         //not used
                    mu.RunDailyWagePercentageMove();
                    break;

                case "daily_move_damage":                       //not used
                    mu.RunDailyMoveDamage();
                    break;

                case "daily_route_wages":                       //not used
                    mu.RunDailyRouteWagesMove();
                    break;

                case "check_job_done":                          //not used
                    string sendPhone = cs.ReturnConfigSetting("NightlyRouteToSlack", "AuditorSendPhone").ToString();
                    string templateID = cs.ReturnConfigSetting("NightlyRouteToSlack", "AuditorTemplateID").ToString();
                    mu.RunCheckJobDoneChecklist(sendPhone, templateID);
                    break;

                case "morning_meeting_time":                    //not used
                    mu.SendMorningMeetingTimeToSlack();
                    break;

                case "check_checklist_move":                    //not used
                    mu.RunCheckChecklists();
                    break;

                case "nightly_checklists_move":                 //not used
                    uc.RunNightlyChecklists("ym", cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_movechecklists").ToString().Replace("\\\\", "\\"));
                    break;

                case "download_daily_tips":                     //not used
                    mu.DownloadDailyTips();
                    break;

                case "square_swipe_percentage_move":            //not used
                    mu.MoveSquareSwipePercentage();
                    break;

                case "move_check_uniform_inventory":            //not used
                    mu.MoveCheckUniformInventory();
                    break;

                case "move_overhead":                           //not used
                    mu.MoveReportOutOverheadWages();
                    break;

                case "move_send_sales_center_leads_welcome":    //not used
                    mu.SendSCLeadsAWelcomeText(true);
                    break;

                case "daily_wage":                                  //not used
                    ju.RunDailyWagePercentage();
                    break;

                case "overall_daily_revenue":                       //not used
                    ju.RunOverallDailyRouteRevenue();
                    break;

                case "square_swipe_percentage":                     //not used
                    ju.JunkSquareSwipePercentage();
                    break;

                case "junk_check_uniform_inventory":                //not used
                    ju.JunkCheckUniformInventory();
                    break;

                case "daily_min_dumps":                             //not used
                    ju.RunDailyMinDumps();
                    break;

                case "dump_location":
                    ju.RunDumpLocation();                           //not used
                    break;

                case "check_checklists":                            //not used
                    ju.RunCheckChecklists();
                    break;

                case "nightly_checklists":                          //not used
                    uc.RunNightlyChecklists("gj", cs.ReturnConfigSetting("NightlyRouteToSlack", "slack_gotjunkchecklists").ToString().Replace("\\\\", "\\"));
                    break;




                default:
                    break;

            }

        }

    }
}
