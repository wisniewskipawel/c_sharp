using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;




namespace QubeMapper_CFP {
    class Program {
        //
        static IEnumerable<string> titles = new string[] { "Capt.", "Countess", "Dr", "Lady", "Messrs", "Miss", "Mr", "Mrs", "Ms", "Prof", "Rev", "Revds", "Rt Hon", "Sir" }.OrderByDescending(t => t.Length);
        static IEnumerable<string> removeNameParts = new string[] { "and", "&" };

        static void Main(string[] args)
        {


            var db = new EstateCraftEntities();


            var sourceFilePath = ConfigurationManager.AppSettings["InputPath"];
            var destFilePath = ConfigurationManager.AppSettings["OutputPath"];

            var allLandlords = db.Applicants.Where(lords => lords.app_id != null && lords.app_dbtype == "L").ToList();       //.ToDictionary(lords => lords.app_key)     
            var allApplicants = db.Applicants.Where(app => app.app_id != null && app.app_dbtype == "A").ToList();    //.ToDictionary(app => app.app_key);
            var allDairess = db.Diaries.Where(d => d.prim_key != null).Take(100).ToList();
            var allProperties = db.Properties.Where(propertis => propertis.prop_key != null).ToList();//.ToDictionary(propertis => propertis.prop_key); 
            var allNegotiator = db.Agents.Where(userl => userl.agent_key != null).ToList();
            var allOfice = db.Offices.Where(ofic => ofic.office_key != null).Take(100).ToList();
           
            myrow(allApplicants,allProperties, allDairess, destFilePath);
           

            //      var allTennants = db.Applicants.Where(tennants => tennants.app_key != null).ToList();
            //      var allOverseas = db.Applicants.GroupBy(c => c.app_oldid).Select(c => c.FirstOrDefault()).ToList();
        }

        static void myrow(List<QubeMapper_CFP.Applicant> allApplicants, List<QubeMapper_CFP.Property> allProperties,List<QubeMapper_CFP.Diary> allDairess, string destFilePath)
        {
            var row = new Dictionary<string, object>();

            //    foreach (var negotiator in allNegotiator) {
            foreach (var applicants in allApplicants) {

                row["id"] = applicants.app_id;
                //     Console.ReadKey();
                row["LAST_CALL_DATE"] = applicants.app_lastcontact.ToString();
                //  Console.WriteLine(row["LAST_CALL_DATE"] = applicants.app_lastcontact);
                row["NEXT_CALL_DATE"] = " ";

                row["START_DATE"] = applicants.app_mort_startdate;
                row["MIN_PRICE"] = applicants.app_minprice;
                row["MAX_PRICE"] = applicants.app_maxprice;
                row["NEGOTIATOR"] = applicants.app_negotiator; // != null ? negotiator.agent_name.ToString() : false.ToString();
                row["TYPE"] = applicants.app_mort_mtype;
                //   Console.WriteLine(applicants.app_mort_mtype);
                row["APPLICANT1_HOUSE_NAME"] = " ";
                row["APPLICANT1_HOUSE_NUMBER"] = ""; // EMPTY
                row["APPLICANT1_ADDRESS1"] = " ";
                row["APPLICANT1_ADDRESS2"] = " ";
                row["APPLICANT1_ADDRESS3"] = " ";
                row["APPLICANT1_ADDRESS4"] = "";
                row["APPLICANT1_POSTCODE"] = " ";
                row["APPLICANT1_COUNTRY"] = " "; //(landlord.LSRTCD != null) ? FormatSortCode(landlord.LSRTCD) : "";
                row["APPLICANT1_HOME1 "] = " ";// landlord.LBKNAME;
                row["APPLICANT1_WORK1"] = " ";// landlord.LACTNO;
                row["APPLICANT1_MOBILE1"] = " ";// property.PADD1;
                row["APPLICANT1_EMAIL1"] = "";
                row["APPLICANT1_FAX1"] = "";
                row["APPLICANT1_HOME2 "] = "";
                row["APPLICANT1_WORK2"] = "";
                row["APPLICANT1_MOBILE2"] = "";
                row["APPLICANT1_EMAIL2"] = "";
                row["APPLICANT1_FAX2"] = "";
                row["APPLICANT1_HOME3 "] = "";
                row["APPLICANT1_WORK3"] = "";
                row["APPLICANT1_MOBILE3"] = "";
                row["APPLICANT1_EMAIL3"] = "";
                row["APPLICANT1_FAX3"] = "";
                row["APPLICANT1_HOME4 "] = "";
                row["APPLICANT1_WORK4"] = "";
                row["APPLICANT1_MOBILE4"] = "";
                row["APPLICANT1_EMAIL4"] = "";
                row["APPLICANT1_FAX4"] = "";
                row["APPLICANT1_HOME5 "] = "";
                row["APPLICANT1_WORK5"] = "";
                row["APPLICANT1_MOBILE5"] = "";
                row["APPLICANT1_EMAIL5"] = "";
                row["APPLICANT1_FAX5"] = "";
                row["APPLICANT2_HOUSE_NAME "] = "";
                row["APPLICANT2_HOUSE_NUMBER"] = "";
                row["APPLICANT2_ADDRESS1"] = "";
                row["APPLICANT2_ADDRESS2 "] = "";
                row["APPLICANT2_ADDRESS3"] = "";
                row["APPLICANT2_ADDRESS4"] = "";
                row["APPLICANT2_POSTCODE"] = "";
                row["APPLICANT2_COUNTRY"] = "";
                row["APPLICANT2_HOME1 "] = "";
                row["APPLICANT2_WORK1"] = "";
                row["APPLICANT2_WORK1"] = "";
                row["APPLICANT2_MOBILE1"] = "";
                row["APPLICANT2_EMAIL1"] = "";
                row["APPLICANT2_FAX1 "] = "";
                row["APPLICANT2_HOME2"] = "";
                row["APPLICANT2_WORK2"] = "";
                row["APPLICANT2_MOBILE2 "] = "";
                row["APPLICANT2_EMAIL2"] = "";
                row["APPLICANT2_FAX2"] = "";
                row["APPLICANT2_HOME3"] = "";
                row["APPLICANT2_WORK3"] = "";
                row["APPLICANT2_MOBILE3 "] = "";
                row["APPLICANT2_EMAIL3"] = "";
                row["APPLICANT2_FAX3"] = "";
                row["APPLICANT2_HOME4 "] = "";
                row["APPLICANT2_WORK4"] = "";
                row["APPLICANT2_MOBILE4"] = "";
                row["APPLICANT2_EMAIL4"] = "";
                row["APPLICANT2_FAX4"] = "";
                row["APPLICANT2_HOME5 "] = "";
                row["APPLICANT2_WORK5"] = "";
                row["APPLICANT2_MOBILE5"] = "";
                row["APPLICANT2_EMAIL5 "] = "";
                row["APPLICANT2_FAX5"] = "";
                row["COMPANY_NAME1"] = applicants.app_iscompany;
                row["TITLE1"] = "";
                row["INITIALS1"] = "";
                row["SURNAME1 "] = "";
                row["COMPANY_NAME2"] = "";
                row["TITLE2"] = "";
                row["INITIALS2 "] = "";
                row["SURNAME2"] = "";
                row["RENT_FREQUENCY"] = "";
                row["TOTALNUM"] = "";
                row["TOTALNUMTO"] = "";
                row["NUM1 "] = "";
                row["NUM1TO"] = "";
                row["NUM2"] = "";
                row["NUM2TO "] = "";
                row["NUM3"] = "";
                row["NUM3TO "] = "";
                row["NUM4"] = "";
                row["NUM4TO"] = "";
                row["MIN_FEET"] = "";
                row["MAX_FEET"] = "";
                row["REQUIREMENT_TYPE "] = "";
                row["REQUIREMENT_STYLE"] = "";
                row["REQUIREMENT_SITUATION"] = "";
                row["REQUIREMENT_AGE "] = "";
                row["REQUIREMENT_LOCATION"] = "";
                row["REQUIREMENT_PARKING "] = "";
                row["REQUIREMENT_SPECIAL"] = "";
                row["REQUIREMENT_FURNISHED"] = "";
                row["REQUIREMENT_DECORATION"] = "";
                row["AREAS"] = applicants.app_minarea;
                row["SOURCE "] = "";
                row["STATUS"] = "";
                row["KEYWORDS"] = "";
                row["SELL_POSITION"] = "";
                row["SELL_STATUS"] = "";
                row["BUY_POSITION "] = "";
                row["BUY_REASON"] = "";
                row["ACTIVE"] = "";
                row["RETAIN"] = "";
                row["SHARED_NEGOTIATORS"] = "";
                row["SHARED_OFFICES "] = "";
                row["TENURE"] = "";
                row["MIN_LEASE"] = "";
                row["MAX_LEASE "] = "";
                row["ACTIVE_FROM_DATE"] = "";
                row["NOTES1 "] = "";
                row["NOTES2"] = "";
                row["NOTES3"] = "";
                row["NOTES4"] = "";
                row["NOTES5"] = " ";
                row["NOTES6 "] = "";
                row["NOTES7"] = "";
                row["NOTES8"] = "";
                row["NOTES9"] = "";
                row["NOTES10"] = "";
                row["APPLICANT1_MARKETING"] = "";
                row["APPLICANT2_MARKETING"] = "";

            }
            Console.WriteLine("Aplikants");
            createfiles(row, destFilePath);
            negotiator(allApplicants, allProperties, allDairess, destFilePath);
        }
        static void negotiator(List<QubeMapper_CFP.Applicant> allApplicants,List<QubeMapper_CFP.Property> allProperties,List<QubeMapper_CFP.Diary> allDairess, string destFilePath)
        {

            var row = new Dictionary<string, object>();
            //           foreach (var negotiator in allNegotiator) {
            foreach (var applicants in allApplicants) {

                //                  row["id"] = applicants.app_id;
                //     Console.ReadKey();
                //                row["LAST_CALL_DATE"] = applicants.app_lastcontact.ToString();
                //  Console.WriteLine(row["LAST_CALL_DATE"] = applicants.app_lastcontact);
                row["NEXT_CALL_DATE"] = " ";

                row["START_DATE"] = applicants.app_mort_startdate;
                row["MIN_PRICE"] = applicants.app_minprice;
                row["MAX_PRICE"] = applicants.app_maxprice;
                row["NEGOTIATOR"] = applicants.app_negotiator;//!=null ? negotiator.agent_name.ToString() : false.ToString();
                row["TYPE"] = applicants.app_mort_mtype;
                //   Console.WriteLine(applicants.app_mort_mtype);
                row["APPLICANT1_HOUSE_NAME"] = " ";
                row["APPLICANT1_HOUSE_NUMBER"] = ""; // EMPTY
                row["APPLICANT1_ADDRESS1"] = " ";
                row["APPLICANT1_ADDRESS2"] = " ";
                row["APPLICANT1_ADDRESS3"] = " ";
                row["APPLICANT1_ADDRESS4"] = "";
                row["APPLICANT1_POSTCODE"] = " ";
                row["APPLICANT1_COUNTRY"] = " "; //(landlord.LSRTCD != null) ? FormatSortCode(landlord.LSRTCD) : "";
                row["APPLICANT1_HOME1 "] = " ";// landlord.LBKNAME;
                row["APPLICANT1_WORK1"] = " ";// landlord.LACTNO;
                row["APPLICANT1_MOBILE1"] = " ";// property.PADD1;
                row["APPLICANT1_EMAIL1"] = "";
                row["APPLICANT1_FAX1"] = "";
                row["APPLICANT1_HOME2 "] = "";
                row["APPLICANT1_WORK2"] = "";
                row["APPLICANT1_MOBILE2"] = "";
                row["APPLICANT1_EMAIL2"] = "";
                row["APPLICANT1_FAX2"] = "";
                row["APPLICANT1_HOME3 "] = "";
                row["APPLICANT1_WORK3"] = "";
                row["APPLICANT1_MOBILE3"] = "";
                row["APPLICANT1_EMAIL3"] = "";
                row["APPLICANT1_FAX3"] = "";
                row["APPLICANT1_HOME4 "] = "";
                row["APPLICANT1_WORK4"] = "";
                row["APPLICANT1_MOBILE4"] = "";
                row["APPLICANT1_EMAIL4"] = "";
                row["APPLICANT1_FAX4"] = "";
                row["APPLICANT1_HOME5 "] = "";
                row["APPLICANT1_WORK5"] = "";
                row["APPLICANT1_MOBILE5"] = "";
                row["APPLICANT1_EMAIL5"] = "";
                row["APPLICANT1_FAX5"] = "";
                row["APPLICANT2_HOUSE_NAME "] = "";
                row["APPLICANT2_HOUSE_NUMBER"] = "";
                row["APPLICANT2_ADDRESS1"] = "";
                row["APPLICANT2_ADDRESS2 "] = "";
                row["APPLICANT2_ADDRESS3"] = "";
                row["APPLICANT2_ADDRESS4"] = "";
                row["APPLICANT2_POSTCODE"] = "";
                row["APPLICANT2_COUNTRY"] = "";
                row["APPLICANT2_HOME1 "] = "";
                row["APPLICANT2_WORK1"] = "";
                row["APPLICANT2_WORK1"] = "";
                row["APPLICANT2_MOBILE1"] = "";
                row["APPLICANT2_EMAIL1"] = "";
                row["APPLICANT2_FAX1 "] = "";
                row["APPLICANT2_HOME2"] = "";
                row["APPLICANT2_WORK2"] = "";
                row["APPLICANT2_MOBILE2 "] = "";
                row["APPLICANT2_EMAIL2"] = "";
                row["APPLICANT2_FAX2"] = "";
                row["APPLICANT2_HOME3"] = "";
                row["APPLICANT2_WORK3"] = "";
                row["APPLICANT2_MOBILE3 "] = "";
                row["APPLICANT2_EMAIL3"] = "";
                row["APPLICANT2_FAX3"] = "";
                row["APPLICANT2_HOME4 "] = "";
                row["APPLICANT2_WORK4"] = "";
                row["APPLICANT2_MOBILE4"] = "";
                row["APPLICANT2_EMAIL4"] = "";
                row["APPLICANT2_FAX4"] = "";
                row["APPLICANT2_HOME5 "] = "";
                row["APPLICANT2_WORK5"] = "";
                row["APPLICANT2_MOBILE5"] = "";
                row["APPLICANT2_EMAIL5 "] = "";
                row["APPLICANT2_FAX5"] = "";
                row["COMPANY_NAME1"] = applicants.app_iscompany;
                row["TITLE1"] = "";
                row["INITIALS1"] = "";
                row["SURNAME1 "] = "";
                row["COMPANY_NAME2"] = "";
                row["TITLE2"] = "";
                row["INITIALS2 "] = "";
                row["SURNAME2"] = "";
                row["RENT_FREQUENCY"] = "";
                row["TOTALNUM"] = "";
                row["TOTALNUMTO"] = "";
                row["NUM1 "] = "";
                row["NUM1TO"] = "";
                row["NUM2"] = "";
                row["NUM2TO "] = "";
                row["NUM3"] = "";
                row["NUM3TO "] = "";
                row["NUM4"] = "";
                row["NUM4TO"] = "";
                row["MIN_FEET"] = "";
                row["MAX_FEET"] = "";
                row["REQUIREMENT_TYPE "] = "";
                row["REQUIREMENT_STYLE"] = "";
                row["REQUIREMENT_SITUATION"] = "";
                row["REQUIREMENT_AGE "] = "";
                row["REQUIREMENT_LOCATION"] = "";
                row["REQUIREMENT_PARKING "] = "";
                row["REQUIREMENT_SPECIAL"] = "";
                row["REQUIREMENT_FURNISHED"] = "";
                row["REQUIREMENT_DECORATION"] = "";
                row["AREAS"] = applicants.app_minarea;
                row["SOURCE "] = "";
                row["STATUS"] = "";
                row["KEYWORDS"] = "";
                row["SELL_POSITION"] = "";
                row["SELL_STATUS"] = "";
                row["BUY_POSITION "] = "";
                row["BUY_REASON"] = "";
                row["ACTIVE"] = "";
                row["RETAIN"] = "";
                row["SHARED_NEGOTIATORS"] = "";
                row["SHARED_OFFICES "] = "";
                row["TENURE"] = "";
                row["MIN_LEASE"] = "";
                row["MAX_LEASE "] = "";
                row["ACTIVE_FROM_DATE"] = "";
                row["NOTES1 "] = "";
                row["NOTES2"] = "";
                row["NOTES3"] = "";
                row["NOTES4"] = "";
                row["NOTES5"] = " ";
                row["NOTES6 "] = "";
                row["NOTES7"] = "";
                row["NOTES8"] = "";
                row["NOTES9"] = "";
                row["NOTES10"] = "";
                row["APPLICANT1_MARKETING"] = "";
                row["APPLICANT2_MARKETING"] = "";
            }
            Console.WriteLine("negotiator");
            createfiles(row, destFilePath);
            property(allProperties, allDairess, destFilePath);
        }

        static void property(List<QubeMapper_CFP.Property> allProperties,List<QubeMapper_CFP.Diary> allDairess, string destFilePath)
        {
            var row = new Dictionary<string, object>();
            foreach (var property in allProperties) {

                // Property
                row["REFERENCE"] = " ";  //.PADD1;
                row["TYPE"] = " ";
                row["DEPARTMENT"] = "";
                row["SITE_TYPE"] = " ";
                row["HOUSE_NAME"] = " "; // (property.PADD4);
                row["ADDRESS1"] = " ";
                row["ADDRESS2"] = " ";
                row["ADDRESS3"] = "";
                row["ADDRESS4"] = "";
                row["prop_postcode"] = " ";
                row["COUNTRY"] = " ";
                row["REFERENCE"] = " ";  //.PADD1;
                row["TYPE"] = " ";
                row["DEPARTMENT"] = "";
                row["SITE_TYPE"] = " ";
                row["HOUSE_NAME"] = " "; // (property.PADD4);
                row["ADDRESS1"] = " ";
                row["ADDRESS2"] = " ";
                row["ADDRESS3"] = "";
                row["ADDRESS4"] = "";
                row["prop_postcode"] = " ";
                row["COUNTRY"] = " ";
                row["AREA"] = property.prop_floorarea;
                row["STRAPLINE"] = " ";
                row["BRIEF_DESCRIPTION "] = "";
                row["NEGOTIATOR"] = " ";
                row["OFFICE"] = " "; // (property.PADD4);
                row["OFFICE2"] = " ";
                row["OFFICE3"] = " ";
                row["REGISTER_DATE"] = "";
                row["LAST_CALL_DATE"] = "";
                row["NEXT_CALL_DATE"] = " ";
                row["EXTERNAL"] = " ";
                row["BOARD_STATUS"] = property.prop_status;  //.PADD1;
                row["BOARD_TYPE"] = " ";
                row["BOARD_DATE"] = "";
                row["URL"] = " ";
                row["URL_TEXT"] = " "; // (property.PADD4);
                row["LATITUDE"] = " ";
                row["LONGITUDE"] = " ";
                row["NO_INTERNET_ADVERTISING"] = "";
                row["TAXBAND"] = "";
                row["LONG_DESCRIPTION"] = " ";
                row["EER"] = " ";
                row["EIR"] = " ";
                row["EIR_POTENTIAL"] = " ";
                row["ARCHIVE_DATE "] = "";
                row["AVAILABLE_FROM_DATE"] = " ";
                row["MARKET_APPRAISAL_DATE"] = " "; // (property.PADD4);
                row["MARKET_APPRAISAL_NEGOTIATOR"] = " ";
                row["KEY_NUMBER"] = " ";
                row["KEY_OFFICE"] = "";
                row["VACANT"] = "";
                row["VIEWING"] = " ";
                row["DISPOSAL"] = " ";
                row["PRICE"] = property.prop_offerprice;
                row["FOR_SALE_DATE"] = "";
                row["EXCHANGE_DATE"] = " ";
                row["COMPLETION_DATE"] = " "; // (property.PADD4);
                row["AGREEMENT_EXPIRY_DATE"] = " ";
                row["PRICE_QUALIFIER"] = " ";
                row["AGENCY"] = "";
                row["SALE_STATUS"] = "";
                row["COMMISSION"] = " ";
                row["EXCHANGE_PRICE"] = " ";
                row["VALUATION_PRICE"] = " ";
                row["ASKING_PRICE"] = " ";
                row["DEVELOPMENT_TOTAL "] = "";
                row["RESERVATION_FEE"] = " ";
                row["RESERVATION_FEE_RETAINED"] = " "; // (property.PADD4);
                row["SERVICE_CHARGE"] = property.prop_servicecharge;
                row["GROUND_RENT"] = " ";
                row["SERVICE_CHARGE_NOTE"] = "";
                row["GROUND_RENT_NOTE"] = "";
                row["EXCHANGE_OFFICE"] = " ";
                row["FEE_NOTE"] = " ";
                row["INTER_OFFICE_FEE"] = " "; // (property.PADD4);
                row["SUBAGENT"] = " ";
                row["SUBAGENT_FEE"] = " ";
                row["JOINT_AGENT_FEE"] = "";
                row["RENT_FREQUENCY"] = "";
                row["RENT_COLLECTION_FEE"] = " ";
                row["FEE_COLLECTION_FREQUENCY"] = " ";
                row["FEE_START_DATE"] = " ";
                row["FEE_END_DATE"] = "";
                row["RENT_INCLUDES"] = " ";
                row["DEPOSIT"] = " "; // (property.PADD4);
                row["DEPOSIT_TYPE"] = " ";
                row["LET_STATUS"] = " ";
                row["FOR_LET_DATE"] = "";
                row["AVAILABLE_TO_DATE"] = "";
                row["ROLE"] = " ";
                row["HAS_GAS"] = " ";
                row["HAS_ELECTRIC_CERT"] = " ";
                row["TOTALNUM"] = " ";
                row["TOTALNUMTO "] = "";
                row["NUM1"] = " ";
                row["NUM1TO"] = " "; // (property.PADD4);
                row["NUM2"] = property.prop_servicecharge;
                row["NUM2TO"] = " ";
                row["NUM3"] = "";
                row["NUM3TO"] = "";
                row["NUM4"] = " ";
                row["NUM4TO"] = " ";
                row["MIN_FEET"] = " "; // (property.PADD4);
                row["MAX_FEET"] = " ";
                row["PLOTS"] = " ";
                row["ACRES"] = "";
                row["ATTRIBUTE_TYPE"] = "";
                row["ATTRIBUTE_STYLE"] = " ";
                row["ATTRIBUTE_SITUATION"] = " ";
                row["ATTRIBUTE_AGE"] = " ";
                row["ATTRIBUTE_LOCATION "] = "";
                row["ATTRIBUTE_PARKING"] = " ";
                row["ATTRIBUTE_SPECIAL"] = " ";
                row["ATTRIBUTE_FURNISHED"] = property.prop_furn;
              //  Console.WriteLine(property.prop_furn);
                row["ATTRIBUTE_DECORATION"] = " ";
                row["KEYWORDS"] = "";
                row["TENURE"] = "";
                row["SALE_TENURE"] = " ";
                row["LET_TENURE"] = " ";
                row["TENURE_END_DATE"] = " ";
                row["MINIMUM_TERM"] = property.prop_leaseterm;
                row["COMPANY_NAME1 "] = "";
                row["TITLE1"] = " ";
                row["INITIALS1"] = " "; // (property.PADD4);
                row["SURNAME1"] = " ";
                row["COMPANY_NAME2"] = " ";
                row["TITLE2"] = "";
                row["INITIALS2"] = "";
                row["SURNAME2"] = " ";
                row["VENDOR1_HOUSE_NAME"] = " ";
                row["VENDOR1_HOUSE_NUMBER"] = " "; // (property.PADD4);
                row["VENDOR1_ADDRESS1"] = " ";
                row["VENDOR1_ADDRESS2"] = " ";
                row["VENDOR1_ADDRESS3"] = "";
                row["VENDOR1_ADDRESS4"] = "";
                row["VENDOR1_POSTCODE"] = " ";
                row["VENDOR1_COUNTRY"] = " ";
                row["VENDOR1_HOME1"] = " ";
                row["VENDOR1_WORK1 "] = "";
                row["VENDOR1_MOBILE1"] = " ";
                row["VENDOR1_EMAIL1"] = " ";
                row["VENDOR1_FAX1"] = " ";
                row["VENDOR1_HOME2"] = " ";
                row["VENDOR1_WORK2"] = "";
                row["VENDOR1_MOBILE2"] = "";
                row["VENDOR1_EMAIL2"] = " ";
                row["VENDOR1_FAX2"] = " ";
                row["VENDOR1_HOME3"] = " ";
                row["VENDOR1_WORK3"] = " ";
                row["VENDOR1_MOBILE3 "] = "";
                row["VENDOR1_EMAIL3"] = " ";
                row["VENDOR1_FAX3"] = " "; // (property.PADD4);
                row["VENDOR1_HOME4"] = " ";
                row["VENDOR1_WORK4"] = " ";
                row["VENDOR1_MOBILE4"] = "";
                row["VENDOR1_EMAIL4"] = "";
                row["VENDOR1_FAX4"] = " ";
                row["VENDOR1_HOME5"] = " ";
                row["VENDOR1_WORK5 "] = "";
                row["VENDOR1_MOBILE5"] = " ";
                row["VENDOR1_EMAIL5"] = " ";
                row["VENDOR1_FAX5"] = " ";
                row["VENDOR2_HOUSE_NAME"] = " ";
                row["VENDOR2_HOUSE_NUMBER"] = "";
                row["VENDOR2_ADDRESS1"] = "";
                row["VENDOR2_ADDRESS2"] = " ";
                row["VENDOR2_ADDRESS3"] = " ";
                row["VENDOR2_ADDRESS4"] = " ";
                row["VENDOR1_WORK3"] = " ";
                row["VENDOR2_POSTCODE "] = "";
                row["VENDOR2_COUNTRY"] = " ";
                row["VENDOR2_HOME1"] = " "; // (property.PADD4);
                row["VENDOR2_WORK1"] = " ";
                row["VENDOR2_MOBILE1"] = " ";
                row["VENDOR2_FAX1"] = "";
                row["VENDOR2_HOME2"] = "";
                row["VENDOR2_WORK2"] = " ";
                row["VENDOR2_MOBILE2"] = " ";
                row["VENDOR2_EMAIL2"] = "";
                row["VENDOR2_FAX2"] = " ";
                row["VENDOR2_HOME3"] = " ";
                row["VENDOR2_WORK3"] = " ";
                row["VENDOR2_MOBILE3"] = " ";
                row["VENDOR2_EMAIL3"] = "";
                row["VENDOR2_FAX3"] = "";
                row["VENDOR2_HOME4"] = " ";
                row["VENDOR2_WORK4"] = " ";
                row["VENDOR2_MOBILE4"] = " ";
                row["VENDOR2_EMAIL4"] = " ";
                row["VENDOR2_FAX4 "] = "";
                row["VENDOR2_HOME5"] = " ";
                row["VENDOR2_WORK5"] = " "; // (property.PADD4);
                row["VENDOR2_MOBILE5"] = " ";
                row["VENDOR2_EMAIL5"] = " ";
                row["VENDOR2_FAX5"] = "";
                row["VENDOR1_MARKETING"] = "";
                row["VENDOR2_MARKETING"] = " ";
                row["SOURCE"] = " ";
                row["NOTES1"] = "";
                row["NOTES2"] = " ";
                row["NOTES3"] = " ";
                row["NOTES4"] = " ";
                row["NOTES5"] = " ";
                row["NOTES6"] = "";
                row["NOTES7"] = "";
                row["NOTES8"] = " ";
                row["NOTES9"] = " ";
                row["NOTES10"] = " ";
                row["CONTACT_OLD_CODE"] = " ";
            }
            Console.WriteLine("property");
            createfiles(row, destFilePath);
            dairess(allDairess, destFilePath);
        }

        static void dairess(List<QubeMapper_CFP.Diary> allDairess, string destFilePath)
        {
            var row = new Dictionary<string, object>();

            foreach (var diar in allDairess) {


                row["REGISTER_DATE"] = diar.d_date;
                row["PROPERTY_REFERENCE"] = " ";
                row["APPLICANT_REFERENCE"] = " ";
                row["MADE_BY"] = "";
                row["OFFICE"] = "";
                row["DATETIME"] = diar.d_date;
                row["TYPE"] = " ";
                row["ENTRY ENTRY "] = diar.d_entrytype;
                row["DURATION "] = " ";
                row["FOLLOWUP_NOTES "] = "";
                row["FOLLOWUP_DATE"] = diar.d_starttime;
      //          Console.WriteLine(diar.d_date);
                row["APPLICANT_CONFIRMED"] = " "; // (property.PADD4);
                row["VENDOR_CONFIRMED"] = " ";
                row["NEGOTIATOR_CONFIRMED"] = " ";
                row["CANCELLED"] = "";
                row["REPEAT_NUMBER"] = "";
                row["REPEAT_UNIT"] = " ";
                row["EXPIRY_DATE"] = " ";
                row["ADDITIONAL_NEGOTIATORS "] = "";
                row["ADDITIONAL_OFFICES"] = " ";
            }
            Console.WriteLine("diar");
            createfiles(row, destFilePath);
            Console.ReadKey();
        }

        static void createfiles(Dictionary<string, object> row, string destFilePath)
        {

            try {

                using (var package = new ExcelPackage(new FileInfo(destFilePath))) {
                    var workbook = package.Workbook;
                    var worksheet = workbook.Worksheets.Add("1");
                    int col = 1;
                    int inputRowNumber = 2;
                    var rowNumber = 1;

                    foreach (var key in row.Keys) {
                        worksheet.Cells[1, rowNumber].Value = key;
                        rowNumber++;
                    }

                    foreach (var values in row.Values) {
                        if (col > row.Keys.Count) {
                            worksheet.Cells[inputRowNumber, col].Value = values;
                            inputRowNumber++;
                            Console.Write(col + "col" + inputRowNumber + "input");
                        }
                        col++;                       
                    }
                    package.SaveAs(new FileInfo(destFilePath));
                }
            }
            catch (Exception e) {
                Console.WriteLine(e.Message);
            }
        }
    }
}

