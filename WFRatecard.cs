using System;
using System.Net; 
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow; 
using System.Net.Http;
using System.Activities;
using System.Text;
using System.IO;
using System.Runtime.Serialization;
using System.Collections.Generic;

namespace WFratecardPlugin
{
    //1. Opvolging na call/registratie
    public class RatecardIncidents : CodeActivity
    {
        //Create fields in the workflow that needs to be filled.
        [RequiredArgument] 
        [Input("API Key")]
        public InArgument<string> APIKey { get; set; }      

        [RequiredArgument]
        [Input("Voornaam")]
        public InArgument<string> inputfirstname { get; set; }

        [RequiredArgument]
        [Input("Achternaam")]
        public InArgument<string> inputlastname { get; set; }

        [RequiredArgument]
        [Input("E-mail")]
        public InArgument<string> Inputemail { get; set; }

        [RequiredArgument]
        [Input("Telefoonnummer")]
        public InArgument<string> inputphonenumber { get; set; }

        [RequiredArgument]
        [Input("Groep Naam")]
        public InArgument<string> inputgroupname { get; set; }

        [RequiredArgument]
        [Input("Eigenaar")]
        public InArgument<string> inputowner { get; set; }

        [RequiredArgument]
        [Input("Klantnaam")]
        public InArgument<string> inputcustomer { get; set; }

        [RequiredArgument]
        [Input("Categorie")]
        public InArgument<string> inputcategory { get; set; }

        [RequiredArgument]
        [Input("Subcategorie")]
        public InArgument<string> inputsubcategory { get; set; }

        [RequiredArgument]
        [Input("Proces")]
        public InArgument<string> inputproces { get; set; }

        [RequiredArgument]
        [Input("Deelproces")]
        public InArgument<string> inputdeelproces { get; set; }

        [RequiredArgument]
        [Input("Bron")]
        public InArgument<string> inputsource { get; set; }

        [RequiredArgument]
        [Input("Afdeling")]
        public InArgument<string> inputdepartment{ get; set; }

        [RequiredArgument]
        [Input("Afgesloten op")]
        public InArgument<string> inputclosedon { get; set; }


        protected override void Execute(CodeActivityContext context )
        {
            try
            {
                //Get CRM tracing service online, used for logging
                ITracingService tracingService = context.GetExtension<ITracingService>();

                //Get API KEY
                string key = APIKey.Get(context);

                //Fill in values for SmartObject
                Smartfields_incidents smartobj = new Smartfields_incidents(); 
                smartobj.owner = inputowner.Get(context);
                smartobj.customer = inputcustomer.Get(context);
                smartobj.category = inputcategory.Get(context);
                smartobj.subcategory = inputsubcategory.Get(context);
                smartobj.proces = inputproces.Get(context);
                smartobj.deelproces = inputdeelproces.Get(context);
                smartobj.source = inputsource.Get(context);
                smartobj.department = inputdepartment.Get(context);
                smartobj.closedon = inputclosedon.Get(context);

                //Get the input from CRM and put them in the NewContact class for JSON serialization
                NewContact_incidents contact = new NewContact_incidents();
                contact.firstname = inputfirstname.Get(context);
                contact.lastname = inputlastname.Get(context); 
                contact.email = Inputemail.Get(context);
                contact.phone = inputphonenumber.Get(context);
                contact.groups.Add(new Groups_incidents() { name = inputgroupname.Get(context), smart_fields = smartobj});

                //Serialize the input from CRM and create a POST request to ratecard
                System.Runtime.Serialization.Json.DataContractJsonSerializer serializer = new System.Runtime.Serialization.Json.DataContractJsonSerializer(contact.GetType());
                MemoryStream ms = new MemoryStream();
                serializer.WriteObject(ms, contact);
                string jsonMsg = Encoding.Default.GetString(ms.ToArray());

                WebRequest request = WebRequest.Create("https://api.ratecard.io/v1/contacts?access_token=" + key); //JrGMbYp9NYgOjLvP0AEpXR5x82eobQ34ODqDlB6ZKz9rGywMaW1d4V7Jm2ZvLV5e;

                request.Method = "POST";
                request.ContentType = "application/json";
                byte[] bytes = System.Text.Encoding.ASCII.GetBytes(jsonMsg.ToString());
                request.ContentLength = bytes.Length;

                System.IO.Stream os = request.GetRequestStream();
                os.Write(bytes, 0, bytes.Length);
                os.Close();
   
                //Write to CRM Log
                tracingService.Trace("Finished! New contact has been added!");

            }
            catch (Exception ex)
            {
                //Get tracing service online, used for logging
                ITracingService tracingService = context.GetExtension<ITracingService>();

                tracingService.Trace("Ratecardplugin: {0}", ex.ToString());
                throw;
            }
        }
    }

    //5. Na Klantbezoek/Contact
    public class RatecardAppointments : CodeActivity
    {
        //Create fields in the workflow that needs to be filled.
        [RequiredArgument]
        [Input("API Key")]
        public InArgument<string> APIKey { get; set; }

        [RequiredArgument]
        [Input("Voornaam")]
        public InArgument<string> inputfirstname { get; set; }

        [RequiredArgument]
        [Input("Achternaam")]
        public InArgument<string> inputlastname { get; set; }

        [RequiredArgument]
        [Input("E-mail")]
        public InArgument<string> Inputemail { get; set; }

        [RequiredArgument]
        [Input("Telefoonnummer")]
        public InArgument<string> inputphonenumber { get; set; }

        [RequiredArgument]
        [Input("Groep Naam")]
        public InArgument<string> inputgroupname { get; set; }

        [RequiredArgument]
        [Input("Eigenaar")]
        public InArgument<string> inputowner { get; set; }

        [RequiredArgument]
        [Input("Klant")]
        public InArgument<string> inputcustomer { get; set; }

        [RequiredArgument]
        [Input("Organisator")]
        public InArgument<string> inputorganiser { get; set; }

        [RequiredArgument]
        [Input("Type")]
        public InArgument<string> inputtype { get; set; }

        [RequiredArgument]
        [Input("Onderwerp")]
        public InArgument<string> inputsubject { get; set; }

        [RequiredArgument]
        [Input("Startdate")]
        public InArgument<string> inputstartdate { get; set; }

        [RequiredArgument]
        [Input("Enddate")]
        public InArgument<string> inputenddate { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                //Get CRM tracing service online, used for logging
                ITracingService tracingService = context.GetExtension<ITracingService>();

                //Get API KEY
                string key = APIKey.Get(context);

                //Fill in values for SmartObject
                Smartfields_appointments smartobj = new Smartfields_appointments();
                smartobj.owner = inputowner.Get(context);
                smartobj.customer = inputcustomer.Get(context);
                smartobj.organiser = inputorganiser.Get(context);
                smartobj.type = inputtype.Get(context);
                smartobj.subject = inputsubject.Get(context);
                smartobj.startdatetime = inputstartdate.Get(context);
                smartobj.enddatetime = inputenddate.Get(context);

                //Get the input from CRM and put them in the NewContact class for JSON serialization
                NewContact_appointments  contact = new NewContact_appointments();

                contact.firstname = inputfirstname.Get(context);
                contact.lastname = inputlastname.Get(context); 
                contact.email = Inputemail.Get(context);
                contact.phone = inputphonenumber.Get(context);
                contact.groups.Add(new Groups_appointments() { name = inputgroupname.Get(context), smart_fields = smartobj });


                //Serialize the input from CRM and create a POST request to ratecard
                System.Runtime.Serialization.Json.DataContractJsonSerializer serializer = new System.Runtime.Serialization.Json.DataContractJsonSerializer(contact.GetType());
                MemoryStream ms = new MemoryStream();
                serializer.WriteObject(ms, contact);
                string jsonMsg = Encoding.Default.GetString(ms.ToArray());

                WebRequest request = WebRequest.Create("https://api.ratecard.io/v1/contacts?access_token=" + key); //JrGMbYp9NYgOjLvP0AEpXR5x82eobQ34ODqDlB6ZKz9rGywMaW1d4V7Jm2ZvLV5e;

                request.Method = "POST";
                request.ContentType = "application/json";
                byte[] bytes = System.Text.Encoding.ASCII.GetBytes(jsonMsg.ToString());
                request.ContentLength = bytes.Length;

                System.IO.Stream os = request.GetRequestStream();
                os.Write(bytes, 0, bytes.Length);
                os.Close();

                //Write to CRM Log
                tracingService.Trace("Finished! New contact has been added!");

            }
            catch (Exception ex)
            {
                //Get tracing service online, used for logging
                ITracingService tracingService = context.GetExtension<ITracingService>();

                tracingService.Trace("Ratecardplugin: {0}", ex.ToString());
                throw;
            }
        }
    }

    //3. Na welkomstcall
    public class RatecardPhonecalls : CodeActivity
    {
        //Create fields in the workflow that needs to be filled.
        [RequiredArgument]
        [Input("API Key")]
        public InArgument<string> APIKey { get; set; }

        [RequiredArgument]
        [Input("Voornaam")]
        public InArgument<string> inputfirstname { get; set; }

        [RequiredArgument]
        [Input("Achternaam")]
        public InArgument<string> inputlastname { get; set; }

        [RequiredArgument]
        [Input("E-mail")]
        public InArgument<string> Inputemail { get; set; }

        [RequiredArgument]
        [Input("Telefoonnummer")]
        public InArgument<string> inputphonenumber { get; set; }

        [RequiredArgument]
        [Input("Groep Naam")]
        public InArgument<string> inputgroupname { get; set; }

        [RequiredArgument]
        [Input("Eigenaar")]
        public InArgument<string> inputowner { get; set; }

        [RequiredArgument]
        [Input("Klant")]
        public InArgument<string> inputcustomer { get; set; }       

        [RequiredArgument]
        [Input("Type")]
        public InArgument<string> inputtype { get; set; }

        [RequiredArgument]
        [Input("Onderwerp")]
        public InArgument<string> inputsubject { get; set; }      

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                //Get CRM tracing service online, used for logging
                ITracingService tracingService = context.GetExtension<ITracingService>();

                //Get API KEY
                string key = APIKey.Get(context);

                //Fill in values for SmartObject
                Smartfields_phonecall smartobj = new Smartfields_phonecall();
                smartobj.owner = inputowner.Get(context);
                smartobj.customer = inputcustomer.Get(context);
                smartobj.type = inputtype.Get(context);
                smartobj.subject = inputsubject.Get(context);

                //Get the input from CRM and put them in the NewContact class for JSON serialization
                NewContact_phonecalls contact = new NewContact_phonecalls();

                contact.firstname = inputfirstname.Get(context);
                contact.lastname = inputlastname.Get(context);
                contact.email = Inputemail.Get(context);
                contact.phone = inputphonenumber.Get(context);
                contact.groups.Add(new Groups_phonecalls() { name = inputgroupname.Get(context), smart_fields = smartobj });


                //Serialize the input from CRM and create a POST request to ratecard
                System.Runtime.Serialization.Json.DataContractJsonSerializer serializer = new System.Runtime.Serialization.Json.DataContractJsonSerializer(contact.GetType());
                MemoryStream ms = new MemoryStream();
                serializer.WriteObject(ms, contact);
                string jsonMsg = Encoding.Default.GetString(ms.ToArray());

                WebRequest request = WebRequest.Create("https://api.ratecard.io/v1/contacts?access_token=" + key); //JrGMbYp9NYgOjLvP0AEpXR5x82eobQ34ODqDlB6ZKz9rGywMaW1d4V7Jm2ZvLV5e;

                request.Method = "POST";
                request.ContentType = "application/json";
                byte[] bytes = System.Text.Encoding.ASCII.GetBytes(jsonMsg.ToString());
                request.ContentLength = bytes.Length;

                System.IO.Stream os = request.GetRequestStream();
                os.Write(bytes, 0, bytes.Length);
                os.Close();

                //Write to CRM Log
                tracingService.Trace("Finished! New contact has been added!");

            }
            catch (Exception ex)
            {
                //Get tracing service online, used for logging
                ITracingService tracingService = context.GetExtension<ITracingService>();

                tracingService.Trace("Ratecardplugin: {0}", ex.ToString());
                throw;
            }
        }
    }

    //1. Overeenkomst
    public class Ratecardleads : CodeActivity
    {
        //Create fields in the workflow that needs to be filled.
        [RequiredArgument]
        [Input("API Key")]
        public InArgument<string> APIKey { get; set; }

        [RequiredArgument]
        [Input("Voornaam")]
        public InArgument<string> inputfirstname { get; set; }

        [RequiredArgument]
        [Input("Achternaam")]
        public InArgument<string> inputlastname { get; set; }

        [RequiredArgument]
        [Input("E-mail")]
        public InArgument<string> Inputemail { get; set; }

        [RequiredArgument]
        [Input("Telefoonnummer")]
        public InArgument<string> inputphonenumber { get; set; }

        [RequiredArgument]
        [Input("Groep Naam")]
        public InArgument<string> inputgroupname { get; set; }

        [RequiredArgument]
        [Input("Salesmanager")]
        public InArgument<string> inputsm { get; set; }

        [RequiredArgument]
        [Input("Klant")]
        public InArgument<string> inputcustomer { get; set; }

        [RequiredArgument]
        [Input("Bron")]
        public InArgument<string> inputorigin { get; set; }

        [RequiredArgument]
        [Input("Functie Contactpersoon")]
        public InArgument<string> inputjobtitle { get; set; }

        [RequiredArgument]
        [Input("Geslacht Contactpersoon")]
        public InArgument<string> inputgender { get; set; }

        [RequiredArgument]
        [Input("Is Intermediair")]
        public InArgument<string> inputisintermediair { get; set; }

        [RequiredArgument]
        [Input("Gemaakt op")]
        public InArgument<string> inputcreatedon { get; set; }

        [RequiredArgument]
        [Input("Startdatum SO")]
        public InArgument<string> inputsostartdate { get; set; }

        [RequiredArgument]
        [Input("Einddatum SO")]
        public InArgument<string> inputsoenddate { get; set; }

        [RequiredArgument]
        [Input("Hoofdbranche")]
        public InArgument<string> inputmainbranche { get; set; }

        [RequiredArgument]
        [Input("Subbranche")]
        public InArgument<string> inputsubbranche { get; set; }

        [RequiredArgument]
        [Input("Heeft E-uur")]
        public InArgument<string> inputhaseuur { get; set; }

        [RequiredArgument]
        [Input("Voorgestelde Diensten")]
        public InArgument<string> inputrecommendedservice { get; set; }

        [RequiredArgument]
        [Input("Gebruikt Contracting")]
        public InArgument<string> inputusescontracting { get; set; }

        [RequiredArgument]
        [Input("Gebruikt flexplein")]
        public InArgument<string> inputusesflexplein { get; set; }

        [RequiredArgument]
        [Input("Gebruikt Easystaffer")]
        public InArgument<string> inputuseseasystaffer { get; set; }


        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                //Get CRM tracing service online, used for logging
                ITracingService tracingService = context.GetExtension<ITracingService>();

                //Get API KEY
                string key = APIKey.Get(context);

                //Fill in values for SmartObject
                Smartfields_leads smartobj = new Smartfields_leads();
                smartobj.salesmanager = inputsm.Get(context);
                smartobj.customer = inputcustomer.Get(context);
                smartobj.origin = inputorigin.Get(context);
                smartobj.jobtitle = inputjobtitle.Get(context);
                smartobj.gender = inputgender.Get(context);
                smartobj.isintermediair = inputisintermediair.Get(context);
                smartobj.createdon = inputcreatedon.Get(context);
                smartobj.sostartdate = inputsostartdate.Get(context);
                smartobj.soenddate = inputsoenddate.Get(context);
                smartobj.mainbranche = inputmainbranche.Get(context);
                smartobj.subbranche = inputsubbranche.Get(context);
                smartobj.haseuur = inputhaseuur.Get(context);
                smartobj.recommendedservices = inputrecommendedservice.Get(context);
                smartobj.usescontracting = inputusescontracting.Get(context);
                smartobj.usesflexplein = inputusesflexplein.Get(context);
                smartobj.useseasystaffer = inputuseseasystaffer.Get(context); 

                //Get the input from CRM and put them in the NewContact class for JSON serialization
                NewContact_leads contact = new NewContact_leads();

                contact.firstname = inputfirstname.Get(context);
                contact.lastname = inputlastname.Get(context); 
                contact.email = Inputemail.Get(context);
                contact.phone = inputphonenumber.Get(context);
                contact.groups.Add(new Groups_leads() { name = inputgroupname.Get(context), smart_fields = smartobj });


                //Serialize the input from CRM and create a POST request to ratecard
                System.Runtime.Serialization.Json.DataContractJsonSerializer serializer = new System.Runtime.Serialization.Json.DataContractJsonSerializer(contact.GetType());
                MemoryStream ms = new MemoryStream();
                serializer.WriteObject(ms, contact);
                string jsonMsg = Encoding.Default.GetString(ms.ToArray());

                WebRequest request = WebRequest.Create("https://api.ratecard.io/v1/contacts?access_token=" + key); //JrGMbYp9NYgOjLvP0AEpXR5x82eobQ34ODqDlB6ZKz9rGywMaW1d4V7Jm2ZvLV5e;

                request.Method = "POST";
                request.ContentType = "application/json";
                byte[] bytes = System.Text.Encoding.ASCII.GetBytes(jsonMsg.ToString());
                request.ContentLength = bytes.Length;

                System.IO.Stream os = request.GetRequestStream();
                os.Write(bytes, 0, bytes.Length);
                os.Close();

                //Write to CRM Log
                tracingService.Trace("Finished! New contact has been added!");

            }
            catch (Exception ex)
            {
                //Get tracing service online, used for logging
                ITracingService tracingService = context.GetExtension<ITracingService>();

                tracingService.Trace("Ratecardplugin: {0}", ex.ToString());
                throw;
            }
        }
    }

    /// <summary>
    /// Classes below are used in conjunction with JSON serialization so it will know where to put the data in the output. 
    /// </summary>
    public class NewContact_incidents
    {
        [DataMember(Name = "first_name")]
        public string firstname { get; set; }
        [DataMember(Name = "last_name")]
        public string lastname { get; set; }
        [DataMember(Name = "email")]
        public string email { get; set; }
        [DataMember(Name = "phone")]
        public string phone { get; set; }
        [DataMember(Name = "groups")]
        public List<Groups_incidents> groups = new List<Groups_incidents>();      
    }

    public class NewContact_appointments
    {
        [DataMember(Name = "first_name")]
        public string firstname { get; set; }
        [DataMember(Name = "last_name")]
        public string lastname { get; set; }
        [DataMember(Name = "email")]
        public string email { get; set; }
        [DataMember(Name = "phone")]
        public string phone { get; set; }
        [DataMember(Name = "groups")]
        public List<Groups_appointments> groups = new List<Groups_appointments>();
    }

    public class NewContact_phonecalls
    {
        [DataMember(Name = "first_name")]
        public string firstname { get; set; }
        [DataMember(Name = "last_name")]
        public string lastname { get; set; }
        [DataMember(Name = "email")]
        public string email { get; set; }
        [DataMember(Name = "phone")]
        public string phone { get; set; }
        [DataMember(Name = "groups")]
        public List<Groups_phonecalls> groups = new List<Groups_phonecalls>();
    }

    public class NewContact_leads
    {
        [DataMember(Name = "first_name")]
        public string firstname { get; set; }
        [DataMember(Name = "last_name")]
        public string lastname { get; set; }
        [DataMember(Name = "email")]
        public string email { get; set; }
        [DataMember(Name = "phone")]
        public string phone { get; set; }
        [DataMember(Name = "groups")]
        public List<Groups_leads> groups = new List<Groups_leads>();
    }

    //Groups Array[object], called in the NewContact class
    public class Groups_incidents
    {
        public string name { get; set; }
        public Smartfields_incidents smart_fields;
    }

    public class Groups_appointments
    {
        public string name { get; set; }
        public Smartfields_appointments smart_fields;

    }

    public class Groups_phonecalls
    {
        public string name { get; set; }
        public Smartfields_phonecall smart_fields;

    }

    public class Groups_leads
    {
        public string name { get; set; }
        public Smartfields_leads smart_fields;

    }


    //Smartfields for incidents entity
    public class Smartfields_incidents
    {
        public string owner { get; set; }
        public string customer { get; set;  }
        public string category { get; set;  }
        public string subcategory { get; set;  }
        public string proces { get; set; }
        public string deelproces { get; set; }
        public string source { get; set;  }
        public string department { get; set; }
        public string closedon { get; set; }
    }

    //Smartfields for appointments entity
    public class Smartfields_appointments
    {
        public string owner { get; set; }
        public string customer { get; set; }
        public string organiser { get; set; }
        public string type { get; set; }
        public string subject { get; set; }
        public string startdatetime { get; set; }
        public string enddatetime { get; set; }
    }

    //Smartfields for phone calls entity
    public class Smartfields_phonecall
    {
        public string owner { get; set; }
        public string type { get; set; }
        public string subject { get; set; }
        public string customer { get; set; }
    }

    //Smartfields for leads entity
    public class Smartfields_leads
    {
        public string salesmanager { get; set; }
        public string customer { get; set; }
        public string origin { get; set; }
        public string jobtitle { get; set; }
        public string gender { get; set; }
        public string isintermediair { get; set; }
        public string createdon { get; set; }
        public string sostartdate { get; set; }
        public string soenddate { get; set; }
        public string mainbranche { get; set; }
        public string subbranche { get; set; }
        public string haseuur { get; set; }
        public string recommendedservices { get; set; }
        public string usescontracting { get; set; }
        public string usesflexplein { get; set; }
        public string useseasystaffer { get; set; } 

    }

}
