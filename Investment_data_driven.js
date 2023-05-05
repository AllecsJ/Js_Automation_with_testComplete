function main (){

  //connects to the excel sheet for data driven testing
  Driver = DDT.ExcelDriver("C:/Users/alexj/Documents/TESTING DOCUMENTS/DATA DRIVE SHEET/data_for_test_cases.xlsx", "iwr_test_cases", true);
    
//logs into moneyline
  login(Driver);
  Delay(20000);
  
  let count = 0;
 
  //LOOP THROUGH THE LINES ROWS OF THE EXCEL SHEET
  while(!Driver.EOF())
  {
  //PROCESS TRANSACTIONS
  personal_investment_to_investment_local(Driver);
  Driver.Next();
  count = count + 1;
  
  //LOG THE NUMBER OF ITERATIONS
  Log.Message("This is iteration: "+count); 
  
  }
}



//setter functions//

function set_username (){
  return Driver.Value(0);
}

function set_browser(){
  //btChrome or btEdge
  return btEdge;
}
//setter functions end//



// external functions //

function login(Driver)
{
  
  let var_username = Driver.Value(0);
 
  Browsers.Item(set_browser()).Navigate("http://mls-env2-st.jmmb.com/moneyline/personal/login.php");
  let browser = Aliases.browser;
  browser.BrowserWindow.Maximize();
  let page = browser.pageLogin;
  let form = page.formWeblogin;
  let textbox = form.querySelector("#txtUser");
  textbox.Click();
  textbox.SetText(var_username);
  textbox.Keys("[Tab]");
  form.passwordboxPassword.SetText(Project.Variables.Password1);
  form.submitbuttonLogin.ClickButton();
  page.Wait();
  form = page.formWebform;
  textbox = form.querySelector("#divQuestions > div > div > input");
  textbox.Click();
  textbox.SetText("ans1");
  form.submitbuttonContinue.ClickButton();
  page.Wait();
  Log.Message("End of Login section");
  
}


function formatDate(date){

// define the initial date
let variable_date = date;

// convert the date string to a Date object
let dt = new Date(variable_date);

// extract the year, month, and day from the Date object
let year = dt.getFullYear();
let month = dt.getMonth() + 1; // Note: January is 0
let day = dt.getDate();

// format the date as a string in the desired format
let formatted_date = year + "-" + ("0" + month).slice(-2) + "-" + ("0" + day).slice(-2);

  return formatted_date;
  Log.Message(formatted_date);
}

    function personal_investment_to_investment_local(Driver) 
    {
       // data driven varibales
      let var_browser = set_browser;
      let var_transacton_type = Driver.Value(1);  //index 0-6 / 7-8 to buy and sell stocks || Default:  Transfer between JMMB accounts'
      let var_source_account = Driver.Value(2); //5 
      let var_payee_account = Driver.Value(3); // 3 - can be index value based on the drop down or "string value surrounded by double quotes"
      let var_payment_type=Driver.Value(4); // FIXED AMOUNT / AVALIABLE BALANCE
      let var_currency=Driver.Value(5); //0-2 // JA$ US$
      let var_amount=Driver.Value(6); // amount of money to be transferred
      let var_is_future_dated = Driver.Value(7); 
      let var_is_recurring = Driver.Value(8); 
      let var_start_date = Driver.Value(9);
      let var_end_date = Driver.Value(10);
      let var_is_next_day = Driver.Value(11);
      let var_get_notified = Driver.Value(12);
      let var_email = Driver.Value(13);
      let var_phone_number = Driver.Value(14);
      let var_add_notes = Driver.Value(15);
      let var_note = Driver.Value(16); //note
      let var_frequency = Driver.Value(17);
      let var_number_of_days = Driver.Value(18);
     // let var_delivery_speed = Driver.Value(22); // "ACH"
  
      //varibalees for new payees
      let var_payee_name = Driver.Value(19);
      let var_payee_alias = Driver.Value(20);
      let var_payee_account_info = Driver.Value(21);
      
      

      
     //selecting transaction type
      Browsers.Item(set_browser()).Navigate("http://mls-env2-st.jmmb.com/moneyline/personal/app/trans/trans.php");
      let browser = Aliases.browser;
      browser.BrowserWindow.Maximize();
      let vselect = browser.pageTrans.formFrmtrans;
      let locator = browser.pageTrans.formFrmtrans;
      let document = browser.pageTrans.formFrmtrans;
      let page = browser.pageTrans.formFrmtrans;
      
      
      
      //format dates
      Log.Message(var_start_date);
      var_start_date = formatDate(var_start_date);
      var_end_date = formatDate(var_end_date);
      Log.Message(var_start_date);

      
   //remove no available balance popup
//      if (document.querySelector("#exposeMask").hasAttribute("block")){
//        page.buttonRaxMsgOk.ClickButton();
//      }
  
      //document.Wait();
  
    //select transaction type
    Log.Message(var_transacton_type);
      vselect.querySelector("#transExpertType").ClickItem(var_transacton_type);
      Log.Message(var_transacton_type);
      page.Wait();
      Log.Message("SELECT TRANSACTION TYPE");
      

    //select source account
      Aliases.browser.pageTrans.formFrmtrans.selectFromWhichAccountWouldYou.ClickItem(var_source_account);
  
      //remove no available balance popup
//      if (document.querySelector("#exposeMask").hasAttribute("block")){
//        page.buttonRaxMsgOk.ClickButton();
//      }
  

    //payee
      browser = Aliases.browser;
      browser.BrowserWindow.Maximize();
      browser.pageTrans.formFrmtrans.selectToWhomWillFundsBePaid.ClickItem(var_payee_account);
      browser.pageTrans.formFrmtrans.selectToWhomWillFundsBePaid.ClickItem(var_payee_account);
    Log.Message("SELECT PAYEE");
  

      //If client is choosing to enter payee details
       //1 is used because of the list in the index
      if(var_transacton_type == 0 && var_payee_account == 1){
        add_new_internal_payee();
      }
      if(var_transacton_type == 1 && var_payee_account == 1){ 
     
          add_new_local_payee();
      }
      if (var_transacton_type == 2 && var_payee_account == 1){
        add_new_international_payee(Driver);
      }  
      Log.Message("NEW PAYEE ADDED");
      
  
//Payment amount
    //payment type
      vselect.selectWhatIsThePaymentType.clickItem(var_payment_type);
      Log.Message("PAYMENT TYPE SELECTED", var_payment_type);
  

    //currency
      vselect.selectCurrency.ClickItem(var_currency);
      Log.Message("CURRENCY SELECTED", var_currency);
      //chrome_RenderWidgetHostHWND.Click(41, 9);
  
    //payment Amount  
    //if payment type is fixed then we will add amount else we ignore the amount box
      if(var_payment_type == "Fixed Amount"){ //ignoring the amount text box if it's available balance
        let textbox = vselect.textboxAmount;
        textbox.Click();
        textbox.SetText("ABC - ~`!@#%^&*()_+-=/.,']\["); //Testing if amount box takes letter
        textbox.SetText(var_amount); //Amount
      }
    Log.Message("PAYMENT AMOUNT SELECTED");
      Delay(1000);
  

      //some accounts are mandatoy future dated. If account is future dated then the checkbox will have attribute disabled.
        if (document.querySelector("#chkShowSchedule").hasAttribute("disabled")){
          var_is_future_dated = true;
        }

    //check if transaction is future dated
    if (var_is_future_dated){
      //check if reccurring
      if (var_is_recurring){
                            let checkbox = browser.pageTrans.formFrmtrans;
                            checkbox.labelDoYouWantToScheduleThis.checkboxChkshowschedule.ClickChecked(true);
                            checkbox.labelRecurringPaymentStanding.radiobuttonRecurringPaymentStand.ClickButton();
                            
                            //frequency
                            checkbox.selectPaymentFrequency.ClickItem(var_frequency);
                          Log.Message("FREQUENCY SET");
                          
                            //start date
                            let textbox = checkbox.textboxStartDateYyyyMmDd;
                            textbox.Click();
                            textbox.SetText(var_start_date);
                          Log.Message("START DATE SET");
                           
                            //end date
                            textbox = checkbox.textboxEndDateYyyyMmDd;
                            textbox.Click();
                            textbox.SetText(var_end_date);
                          Log.Message("END DATE SET");
                        
                            //if every x days is clicked the number of days is set
                            if(var_frequency = "Every (x) days"){
                                let checkbox = browser.pageTrans.formFrmtrans;
                                checkbox.selectPaymentFrequency.ClickItem("Every (x) days");
                                let textbox = checkbox.textboxNumberOfDays;
                                textbox.Click();
                                textbox.SetText("10");
                                Log.Message("ADD FREQUENCY FOR EVERY X DAYS");
                            }
                        
                        
                            //if next business day is true
                            if (var_is_next_day){
                              checkbox.labelNextBusinessDay.radiobuttonNextBusinessDay.ClickButton();
                              Log.Message("IF NEXT BUSINESS DAY IS ADDED");
                            }
                            //if previous day
                            checkbox.labelBusinessdayoptionlabel.radiobuttonPreviousBusinessDay.ClickButton();
                            Log.Message("IF PERVIOUS DAY IS SELECTED");
                      }else{
                       //if not recurring
                            browser = Aliases.browser;
                            let checkbox = browser.pageTrans.formFrmtrans;
                            checkbox.labelDoYouWantToScheduleThis.checkboxChkshowschedule.ClickChecked(true);
                            checkbox.radiobuttonOneTimePayment.ClickButton();
                            let textbox = checkbox.textboxEffectiveDate;
                            textbox.Click();
                            textbox.SetText(var_start_date);
                            checkbox.panelDivscheduleinfo.Click();
                            Log.Message("IF NOT RECURRING");
                      }
  
    }

    //add notified
    if (var_get_notified){
      vselect.labelDoYouWishToBeNotifiedWhen.checkboxDoYouWishToBeNotifiedWhe.ClickChecked(true);
      vselect.emailinputEmail.Click();
      vselect.emailinputEmail.SetText(var_email);
      vselect.textboxPhoneNumber.Click(); 
      vselect.textboxPhoneNumber.SetText(var_phone_number);
      Log.Message("CONTACT INFO ADDED");
    }

    //add notes
    if (var_add_notes){
      vselect.labelDoYouWantToAddANoteTo.checkboxChkspecialnotes.ClickChecked(true);
      let textarea = vselect.textareaPersonalNoteNoteToSelf;
      textarea.Click();
      textarea.Keys(var_note);
      Log.Message("NOTE ADDED");
      Log.Message(var_note);
    }

    //click the continue button
     vselect.submitbuttonContinue.ClickButton();
     Log.Message("TRANSACTION SUBMITTED");
 

      //remove no available balance popup
//      if (document.querySelector("#exposeMask").hasAttribute("block")){
//        page.buttonRaxMsgOk.ClickButton();
//      }

         page.Wait()
         Delay(10000);
           
          browser = Aliases.browser;
          browser.BrowserWindow.Maximize();
          page = browser.pageTrans;
          page.textnodeRaxMsgboxTitle.Click();
          page.textnode.Click();
          page.buttonRaxMsgOk.ClickButton();
 
          //process transaction
            page.Wait();
            Delay(1000);
            
            //continue
          browser = Aliases.browser;
          browser.BrowserWindow.Maximize();
          page = browser.pageTrans;
          page.formFrmtrans.submitbuttonContinue.ClickButton();
            
            if (var_transacton_type == 0 || var_transacton_type == 1){
              process_local_transaction();
            } 
            
            if (var_transacton_type == 2){
            proccess_international_transactions();
          }

          
          
          
}

function checkpoints (){
   Driver = DDT.ExcelDriver("C:/Users/alexj/Documents/TESTING DOCUMENTS/DATA DRIVE SHEET/data_for_test_cases.xlsx", "Investment_test_cases", true);
    let var_amount=Driver.Value(6);
    let var_start_date = Driver.Value(9);
    let var_currency=Driver.Value(5);
  //checkpoints
          aqObject.CheckProperty(Aliases.browser.pageTrans.cell2, "contentText", cmpEqual, var_start_date);
          aqObject.CheckProperty(Aliases.browser.pageTrans.textnodeJa19900, "contentText", cmpEqual, var_currency+" "+var_amount);

}

//Internal Transfer from a JMMB Investment  account to a JMMB Bank Account (FX)


//Internal Transfer from a JMMB Investment account to a JMMB Bank Account (JMD to JMD)


//Internal Transfer from a JMMB Investment  account to a JMMB Bank Account (Cross Currency)

//Internal Transfer One Time Standing Order Investment to Investment 

//Internal Transfer Recurring Standing Order Investment to Investment 

//Internal Transfer Recurring Standing Order Investment to Bank

//Internal Transfer One Time Standing Order Investment to Bank
function process_local_transaction(){
  browser = Aliases.browser;
          browser.BrowserWindow.Maximize();
          page = browser.pageTrans;
          //page wait
          page.Wait();
          Delay(1000);
          //enter pin
          let passwordBox = page.formFrmtrans2;
          passwordBox.passwordboxEnterYourPin.SetText(Project.Variables.Password2);
          //process
          passwordBox.submitbuttonProcessAllTransactio.ClickButton();
          browser = Aliases.browser;
          browser.BrowserWindow.Maximize();
          page = browser.pageTrans;
          page.linkShowDetails.Click();
          //show details
          //copy details
          page.cell.Drag(6, 11, 55, 5);
          let wndChrome_WidgetWin_1 = browser.wndChrome_WidgetWin_1;
          OCR.Recognize(wndChrome_WidgetWin_1).BlockByText("166").ClickR();
          wndChrome_WidgetWin_1.Click(83, 46);
          
//           //checkpoints
//          aqObject.CheckProperty(Aliases.browser.pageTrans.cell2, "contentText", cmpEqual, var_start_date);
//          aqObject.CheckProperty(Aliases.browser.pageTrans.textnodeJa19900, "contentText", cmpEqual, var_currency+" "+var_amount);
}

function proccess_international_transactions()
{
  //add pin
  let browser = Aliases.browser;
  browser.BrowserWindow.Maximize();
  let page = browser.pageTrans;
  let form = page.formFrmtrans;
  let passwordBox = form.passwordboxSecurityPin;
  passwordBox.Click();
  passwordBox.SetText(Project.Variables.Password2);
  form.querySelector("#questTxt1").SetText("ans1");
  get_verification_code();
  
  
}

function get_verification_code()
{
  let explorer = Aliases.explorer;
  explorer.wndShell_TrayWnd.ReBarWindow32.MSTaskSwWClass.MSTaskListWClass.Click(96, 42);
  explorer.wndWorkerW.SHELLDLL_DefView.FolderView.Drag(1647, 251, 7, 19);
  explorer.wndCabinetWClass.Maximize();
  let directUIHWND = explorer.wndCabinetWClass2.ShellTabWindowClass.DUIViewWndClassName.DirectUIHWND;
  directUIHWND.CtrlNotifySink.NamespaceTreeControl.tvNamespaceTreeControl.ClickItem("|Quick access|Documents");
  let directUIHWND2 = directUIHWND.CtrlNotifySink2.ShellView.DirectUIHWND;
  OCR.Recognize(directUIHWND2).BlockByText("DOCUMENTS").DblClick();
  directUIHWND2.DblClick(130, 82);
  OCR.Recognize(directUIHWND2).BlockByText("getVerificationCodes").Click();
  OCR.Recognize(directUIHWND2).BlockByText("getVerification").ClickR();
  directUIHWND2.PopupMenu.Click("Open");
  let wnd = Aliases.Ssms.Item7;
  wnd.Click(297, 98);
  wnd.Click(304, 94);
  let wnd2 = wnd.GenericPane.WindowsForms10Window8app01e4efd9r45ad1.WindowsForms10Window8app01e4efd9r45ad1.WindowsForms10Window8app01e4efd9r45ad12.Item;
  wnd2.ClickTab("Results");
  let windowsForms10Window8app01e4efd9r45ad1 = wnd2.Results.WindowsForms10Window8app01e4efd9r45ad1.WindowsForms10Window8app01e4efd9r45ad1.WindowsForms10Window8app01e4efd9r45ad1;
  windowsForms10Window8app01e4efd9r45ad1.HScroll.Pos = 0;
  windowsForms10Window8app01e4efd9r45ad1.Click(827, 29);
  windowsForms10Window8app01e4efd9r45ad1.Keys("^c");
  let browser = Aliases.browser;
  browser.BrowserWindow.Position(-9, 0, 987, 1050);
  let page = browser.pageTrans;
  let textbox = page.formFrmtrans.textboxSecurityVerificationCode;
  textbox.Click();
  textbox.Keys("^v");
  page.formFrmtrans2.submitbuttonProcessAllTransactio.ClickButton();
}


function verify_investment_transaction(){
  //selecting transaction type
      Browsers.Item(set_browser()).Navigate("http://appsvr4-env2-st:81/cis/account_search_form.php");
      let browser = Aliases.browser;
      browser.BrowserWindow.Maximize();
      let vselect = browser.pageTrans.formFrmtrans;
      let locator = browser.pageTrans.formFrmtrans;
      let document = browser.pageTrans.formFrmtrans;
 
  
}



function add_new_internal_payee(){
  //payee name
          let textbox = vselect.textboxPayeename;
          textbox.Click();
          textbox.Keys("Jackson Industries");
          //payee alias
          textbox = vselect.textboxPayeejmmbaccountalias;
          textbox.Click();
          textbox.SetText("ja industry");
          //select Company
          let vselect2 = vselect.selectPayeecompanies;
          vselect2.ClickItem("JMMB Bank"); //
          textbox = vselect.textboxPayeejmmbaccount;
          textbox.Click();
          //payee account number
          textbox.SetText(var_payee_account_info);
          //account number currency
          browser.pageTrans.formFrmtrans.selectPayeeseljmmbaccurrency.ClickItem(var_currency);
}


function add_new_local_payee(var_payee_name, var_payee_alias, var_payee_bank_name, var_payee_account_info)
{
          //payee name
          let textbox = vselect.textboxPayeename;
          textbox.Click();
          textbox.Keys(var_payee_name);
          //payee alias
          textbox = vselect.textboxPayeejmmbaccountalias;
          textbox.Click();
          textbox.SetText(var_payee_alias);
          //select Company
          let vselect2 = vselect.selectPayeecompanies;
          vselect2.ClickItem(var_payee_bank_name); //
          textbox = vselect.textboxPayeejmmbaccount;
          textbox.Click();
          //payee account number
          textbox.SetText(var_payee_account_info);
          //account number currency
          browser.pageTrans.formFrmtrans.selectPayeeseljmmbaccurrency.ClickItem(var_currency);
}


function add_new_international_payee()
{
  //international accounts variable
      let var_intl_payee = Driver.Value(23);
      let var_intl_alias = Driver.Value(24);
      let var_intl_acct_no = Driver.Value(25);
      let var_intl_routing = Driver.Value(26);
      let var_intl_routing_method = Driver.Value(27);
      let var_intl_currency = Driver.Value(28);
      let var_intl_country = Driver.Value(29);
      let var_intl_bank_name = Driver.Value(30);
      //intermediary bank variables
      let var_use_intermediary_bank = Driver.Value(31);
      let var_intermediary_country = Driver.Value(32);
      let var_intermediary_bank_name = Driver.Value(33);
      let var_intermediary_routing_method = Driver.Value(34);
      let var_intermediary_routing_number = Driver.Value(35);
      
  //enter new pagee details
  let browser = Aliases.browser;
  browser.BrowserWindow.Maximize();
  let vselect = browser.pageTrans.formFrmtrans;
  //to whom will funds be paid?
  let textbox = vselect.textboxPayeename;
  textbox.Click();
  textbox.SetText(var_intl_payee);
  textbox = vselect.textboxPayeejmmbaccountalias;
  textbox.Click();
  textbox.SetText(var_intl_alias);
  textbox = vselect.textboxPayeeinwaccountno;
  textbox.Click(var_intl_routing_method);
  textbox.SetText(var_intl_acct_no);
  //Country details
  //test
  vselect.selectPayeeselinwcurrency.ClickItem("GB");
  //vselect.selectPayeeselinwcurrency.ClickItem(var_intl_currency);
  vselect.querySelector("#payeeselinwbbdCountry").ClickItem(var_intl_country);
  //.ClickItem(var_intl_country);
  textbox = vselect.textboxPayeeinwbbdname;
  textbox.Click();
  textbox.SetText(var_intl_bank_name);
  vselect.selectPayeeselinwbbdrouting.ClickItem(var_intl_routing_method);
  //chrome_RenderWidgetHostHWND.Click(257, 53);
  textbox = vselect.textboxPayeeinwrouting;
  textbox.Click();
  textbox.SetText(var_intl_routing);
  
  //var_use_intermediary_bank,var_intermediary_country, var_intermediary_bank_name, var_intermediary_routing_method, var_intermediary_routing_number
  //intermediary bank details
  if(var_use_intermediary_bank){
  Aliases.explorer.wndShell_TrayWnd.ReBarWindow32.MSTaskSwWClass.MSTaskListWClass.Click(45, 37);
  //select intermeiary bank
  //select country
  browser = Aliases.browser;
  browser.BrowserWindow.Position(-8, 0, 976, 1038);
  let form = browser.pageTrans.formFrmtrans;
  form.selectPayeeselinwibcountry.ClickItem(var_intermediary_country);
  //enter bankname
  form.textboxPayeeinwibname.SetText(var_intermediary_bank_name);
  //routing method
  form.selectPayeeselinwidrouting.Click(181, 18);
  document.querySelector("#payeeselinwbbdRouting").ClickItem(var_intermediary_routing_method);
   //routing number
  textbox = form.textboxPayeeinwibrouting;
  textbox.Click();
  textbox.SetText(var_intermediary_routing_number);
    
  }
}
