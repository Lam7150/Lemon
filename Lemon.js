// Opening Spreadsheets
var MasterList = SpreadsheetApp.openById("1NSXptUb7xLx7dl8M19sUiut3PaeGTsJNlbdlnTdfy9s");
var TeamInfo = SpreadsheetApp.openById("1vG0zlJRyMJHFvi1f8hRBIjFOF7wTN7PllLWfLBPnmFA");
var RulesAndWaivers = SpreadsheetApp.openById("1TxTCjQ67K6bx2_VWOKVxN4a1VK1ypRDkTWhLi-8wB-g");
var Health = SpreadsheetApp.openById("1Hatb3g0_DxudxoZFgSufHuoAKgu4V9_vTv8NJ7cxTZ8");
var ErrorForm = SpreadsheetApp.openById("1e3zJ-_UXJuhvjOSomgsnCcYaN1D-_AGVGsbjcBYIfpI");

function Lemon() {
  // Automatically organizes new Team info from jotform onto Master List spreadsheet
  // Getting Team information
  var newTeamIndex = TeamInfo.getLastRow();                                                     // Index of new Team's information           
  var newTeam = TeamInfo.getRange("B"+newTeamIndex+":CI"+newTeamIndex).getValues()[0];          // Team info
  var CeremonyShirtsIndex = MasterList.getSheetByName("For Opening Ceremony").getLastRow() + 1; // Index of Ceremony/Shirts sheet to input new team
  var MasterListIndex = MasterList.getSheetByName("Master List").getLastRow();                  // Index of Master List to input new team
    if (MasterListIndex == 8)                                                                   // If first team, sets correct index (grade count is counted as LastRow)
      MasterListIndex = 0;
    MasterListIndex += 2;                                                                       // Adding spacing between teams
  
  /***   Master List   ***/
  // Setting Team Name
  MasterList.getRange("A"+MasterListIndex).setValue(newTeam[0]);
    
  // Sanitizing and splitting Team Members into array
  var TeamMembers = [[n.replace(/[^a-z ]/gi, '')] for each (n in String(newTeam[6]).split("\n"))]; 
      
  // Setting Team Members
  MasterList.getRange("B"+MasterListIndex+":B" + (MasterListIndex+TeamMembers.length-1)).setValues(TeamMembers);
  
  // Resolving any errors for members who signed forms before the team signed up
  for (var i = 0; i < TeamMembers.length; i++) {
    checkAndResolveError(TeamMembers[i][0]);
  }
  
  // Getting Coaches
  var Coaches = [];
  for (var n = 14; n < newTeam.length; n+=5){ 
    //Ensures empty coaches are not included
    if (newTeam[n] != "")
      Coaches.push([newTeam[n] + " " + newTeam[n+1]]);
  }
  
  // Setting Team Coaches
  MasterList.getRange("F"+(MasterListIndex)+":F"+(MasterListIndex+Coaches.length-1)).setValues(Coaches);
  
  /***   Opening Ceremony   ***/
  MasterList.getRange("For Opening Ceremony!A"+CeremonyShirtsIndex).setValue(newTeam[0]); // Setting Team Names
  MasterList.getRange("For Opening Ceremony!B"+CeremonyShirtsIndex).setValue(newTeam[7]); // Setting Team Songs 
  MasterList.getRange("For Opening Ceremony!C"+CeremonyShirtsIndex).setValue(newTeam[9]); // Setting Team Descriptions 
  
  /***   Shirts   ***/
  // Setting Team Names 
  MasterList.getRange("Shirts!A"+CeremonyShirtsIndex).setValue(newTeam[0]);
  
  // Cleaning shirt input
  var TeamShirts = [[n.slice(-1) for each (n in String(newTeam[11]).split("\n"))]];
  
  // Setting shirt values
  if (TeamShirts[0] != "") {
    // Formatting shirt values so that blanks appear as "0"
    for (var i = 0; i < TeamShirts[0].length; i++) {
      if (TeamShirts[0][i] == " ")
          TeamShirts[0][i] = "0";
    }
  } 
  else // Teams with no requested shirts
    TeamShirts[0] = ["-","-","-","-","-","-"];
  
  // Setting Shirts 
  MasterList.getRange("Shirts!B"+CeremonyShirtsIndex+":G"+CeremonyShirtsIndex).setValues(TeamShirts);
}


function RulesAndWaiversCheck() {
  // Automatically checks off person who has submitted their Rules and Waivers Form
  // Declaring Variables
  var Members = MasterList.getRange("B2:B"+MasterList.getLastRow()).getValues();                                         // List of people already signed up on relay on Master List
  var Person = RulesAndWaivers.getRange("C"+RulesAndWaivers.getLastRow()+":D"+RulesAndWaivers.getLastRow()).getValues(); // Person who submitted form
      Person = String(Person[0][0] + " " + Person[0][1]);
  var personIndex = LookForName(Person, Members)+2;                                                                        // Searches for person on Master List
  
  // Checking off person on Master List
  if (personIndex >= 2)
    MasterList.getRange("Master List!D"+(personIndex)).setValue('X');
  else {
    // If person is not found on Master List, error will be reported to be fixed manually later
    ReportError(ErrorForm.getLastRow()+1, "Rules and Waivers Form Check: "+Person+" not found in Master List", Person);
  }
  MasterList.getRange("Master List!D2:D").setHorizontalAlignment("center");
}


function HealthCheck() {
  // Automatically checks off person who has submitted their Health Form
  // Declaring Variables
  var Members = MasterList.getRange("B2:B"+MasterList.getLastRow()).getValues();              // List of people already signed up on relay on Master List
  var Person = Health.getRange("C"+Health.getLastRow()+":D"+Health.getLastRow()).getValues(); // Person who submitted form
      Person = String(Person[0][0] + " " + Person[0][1]);
  var Grade = Health.getRange("E"+Health.getLastRow()).getValue();                            // Person's grade
  var personIndex = LookForName(Person, Members)+2;                                           // Searches for person on Master list
  
  // Checking off person on Master List and setting grade
  if (personIndex >= 2) {
    MasterList.getRange("Master List!C"+(personIndex)).setValue(Grade); // Setting grade
    MasterList.getRange("Master List!E"+(personIndex)).setValue('X');   // Checking off person
  }
  else {
    // If person is not found on Master List, error will be reported to be fixed manually later
    ReportError(ErrorForm.getLastRow()+1, "Health Form Check: "+Person+" not found in Master List", Person);
  }
  MasterList.getRange("Master List!C2:E").setHorizontalAlignment("center");
}


function LookForName(name, list) {
  // Searches for name in list (of people already signed up for the relay on Master List)
  for (var n = 0; n < list.length; n++){
    if (name.toUpperCase() == list[n][0].toUpperCase())
      return n;
  }
  return -1;
}


function ReportError(index, errorDesc, personAffected) {
  var timestamp = Utilities.formatDate(new Date(), "GMT-04:00", "MM-dd-yyyy HH:mm:ss");
  var error = [[timestamp, errorDesc, personAffected]];
  ErrorForm.getRange("A"+index+":C"+index).setValues(error);
  ErrorForm.getRange("A"+index).setHorizontalAlignment("left");
}


function checkAndResolveError(person){
  // Checks error form for errors for specified person and corrects them
  var errorList = ErrorForm.getRange("C2:C"+ErrorForm.getLastRow()).getValues();
  var errorFormIndex = LookForName(person, errorList)+2;

  if (errorFormIndex >= 2) {
    var masterListMembers = MasterList.getRange("B2:B"+MasterList.getLastRow()).getValues(); // List of people already signed up on relay on Master List
    var masterListIndex = LookForName(person, masterListMembers)+2;
    var errorType = ErrorForm.getRange("B"+errorFormIndex).getValue().charAt(0);
    
    // Checking off person on Master list
    if (errorType === 'R')                                                    // Rules and Waivers
      MasterList.getRange("Master List!D"+(masterListIndex)).setValue('X');  
    else if (errorType === 'H')                                               // Health Form
      MasterList.getRange("Master List!E"+(masterListIndex)).setValue('X'); 
    
    // Deleting error and checking for other errors
    DeleteError(errorFormIndex);
    checkAndResolveError(person);
  }
}


function DeleteError(index) {
  // Deletes error on the error form and shifts rest of form up
  ErrorForm.getRange("A"+(index+1)+":C").moveTo(ErrorForm.getRange("A"+index+":C"));
}


function SetTriggers() {
  // Setting Triggers for functions to run automatically when new forms are submitted
  ScriptApp.newTrigger("Lemon").forSpreadsheet(TeamInfo).onChange().create();                        // Lemon Trigger
  ScriptApp.newTrigger("RulesAndWaiversCheck").forSpreadsheet(RulesAndWaivers).onChange().create();  // Rules and Waivers Trigger
  ScriptApp.newTrigger("HealthCheck").forSpreadsheet(Health).onChange().create();                    // Health Form Trigger
}


function Clear() {
  /** 
   * Clears all spreadsheets for new year
   * *** DO NOT USE UNTIL NEW YEAR, DELETES ALL DATA WITHOUT UNDO. ***
   * Recommendation: Make copies of spreadsheets and save data from each year before clearing for future developers.
   *                 Copy old spreadsheets into Lemon -> Test Data folder.
   */
  
  MasterList.getRange("Master List!A2:H").clear();           // Clears Master List
  MasterList.getRange("For Opening Ceremony!A2:C").clear();  // Clears Opening Ceremony sheet
  MasterList.getRange("Shirts!A2:G").clear();                // Clears Shirts sheet
  ErrorForm.getRange("A2:C").clear();                        // Clears Error Form
  
  // Because of the fact that it is very likely someone will accidentally clear the forms, this has been commented out until
  // someone deliberately and consciously wants to delete these forms for the next year. SAVE before deleting for test data
  // To uncomment functions below, remove the // marks. To recomment them afterwards, put the // marks back.
  //TeamInfo.getRange("A2:CP").clear();                        // Clears Team Information Form
  //RulesAndWaivers.getRange("A2:Q").clear();                  // Clears Rules and Waivers Form
  //Health.getRange("A2:BH").clear();                          // Clears Health Form
}


function DeleteTriggers() {
  // Programmatically clears all triggers set
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  } 
}


function OrganizeAll() {
  /** 
   *************************************************************
   * FOR DEMONSTRATION PURPOSES ONLY, OVERWRITES ALL DATA ON DOC
   *************************************************************
   * Organizes all the information from 
   * spreadsheets of data from team sign-up forms
   * on jotform onto LIME formatted spreadsheet at once
   * (Currently 2018 Version, takes approximately 30 seconds)
   */
  
  // Getting Team Information 
  var numTeams = TeamInfo.getLastRow() - 1;
  var TeamNames = TeamInfo.getRange("B2:B" + (numTeams + 1)).getValues();
  var TeamSongs = TeamInfo.getRange("I2:I" + (numTeams + 1)).getValues(); 
  var TeamDescriptions = TeamInfo.getRange("K2:K" + (numTeams + 1)).getValues();
  var TeamShirts = TeamInfo.getRange("M2:M" + (numTeams + 1)).getValues();
  var TeamMembers = TeamInfo.getRange("H2:H" + (numTeams + 1)).getValues();
  var Coaches = TeamInfo.getRange("P2:CI" + (numTeams + 1)).getValues();
  
  // Other Variables
  var numCoaches = 0;
  var nextLine = 2;
  
  /***   Master List   ***/
  /* Iterates through each team and records information onto Master List */
  for (var i = 0; i < numTeams; i++) {
    // Setting Team Name
    MasterList.getRange("A" + nextLine).setValue(TeamNames[i]);
    
    // Sanitizing and splitting Team Members into array
    TeamMembers[i] = [[n.replace(/[^a-z ]/gi, '')] for each (n in String(TeamMembers[i]).split("\n"))]; 
    
    // Setting Team Members
    MasterList.getRange("B" + nextLine + ":B" + (nextLine+TeamMembers[i].length-1)).setValues(TeamMembers[i]);
    
    // Sanitizing Coach input
    numCoaches = Coaches[i].length
    for (var n = 0; n < numCoaches; n+=5){
      if (Coaches[i][n] != "") { // Adding coach names to end of array
        Coaches[i].push([Coaches[i][n] + " " + Coaches[i][n+1]]);
      }
    }
    Coaches[i].splice(0, numCoaches); // Removing original part of array to leave only coach names
    
    // Setting Team Coaches
    MasterList.getRange("F"+(nextLine)+":F"+(nextLine+Coaches[i].length-1)).setValues(Coaches[i]);
    
    // Determining next starting line based on whether # of members or # of coaches was greater
    nextLine += Math.max(TeamMembers[i].length, Coaches[i].length) + 1; 
  }
  
  /***   Opening Ceremony   ***/
  MasterList.getRange("For Opening Ceremony!A2:A"+(2+numTeams-1)).setValues(TeamNames);        // Setting Team Names
  MasterList.getRange("For Opening Ceremony!B2:B"+(2+numTeams-1)).setValues(TeamSongs);        // Setting Team Songs 
  MasterList.getRange("For Opening Ceremony!C2:C"+(2+numTeams-1)).setValues(TeamDescriptions); // Setting Team Descriptions 
  
  /***   Shirts   ***/
  // Setting Team Names 
  MasterList.getRange("Shirts!A2:A"+(2+numTeams-1)).setValues(TeamNames);
  
  // Cleaning shirt input
  for (var i = 0; i < numTeams; i++) {
    // Slicing shirts into arrays
    TeamShirts[i] = [n.slice(-1) for each (n in String(TeamShirts[i]).split("\n"))];
    
    // Setting shirt values
    if (TeamShirts[i] != "") {
      // Formatting shirt values so that blanks appear as "0"
      for (var n = 0; n < TeamShirts[i].length; n++) {
        if (TeamShirts[i][n] == " ")
            TeamShirts[i][n] = "0";
      }
    } 
    else // Teams with no requested shirts
      TeamShirts[i] = ["-","-","-","-","-","-"];
  }
  
  // Setting Shirts 
  MasterList.getRange("Shirts!B2:G"+(2+numTeams-1)).setValues(TeamShirts);

  /*** Rules and Waivers ***/
  var Members = MasterList.getRange("B2:B"+MasterList.getLastRow()).getValues();                         // List of people already signed up on relay on Master List 
  var RulesAndWaiversForms = RulesAndWaivers.getRange("C2:D"+RulesAndWaivers.getLastRow()).getValues();  // List of people who've submitted Rules and Waivers forms
  for(var RulesAndWaiversConfirmed = []; RulesAndWaiversConfirmed.length < Members.length; RulesAndWaiversConfirmed.push([" "])); // Setting up an empty array to track people checked on Master List
  var Person, personIndex, personGrade; 
  
  // Checking off each person on Master List who've submitted their Rules and Waivers form
  for (var i = 0; i < RulesAndWaiversForms.length; ++i) {
    // Getting person's name and searching for name on Master List
    Person = String(RulesAndWaiversForms[i][0] + " " + RulesAndWaiversForms[i][1]);
    personIndex = LookForName(Person, Members);
    
    // Checking off person's name on Master List
    if (personIndex >= 0)
      RulesAndWaiversConfirmed[personIndex] = ['X'];
    else
      ReportError(ErrorForm.getLastRow()+1, "Rules and Waivers Check: "+Person+" not found in Master List", Person);
  }
  
  // Setting 'X's on each person who've submitted their Rules and Waivers forms on Master List
  MasterList.getRange("Master List!D2:D"+(2+Members.length-1)).setValues(RulesAndWaiversConfirmed);
  
  /*** Health Forms ***/
  var HealthForms = Health.getRange("C2:E"+Health.getLastRow()).getValues();   // List of people who've submitted their Health forms
  for(var HealthConfirmed = []; HealthConfirmed.length < Members.length; HealthConfirmed.push([" "])); // Setting up an empty array to track people checked on Master List
  for(var grades = []; grades.length < Members.length; grades.push([" "])); // Setting up an empty array to track people checked on Master List
  
  // Checking off each person on Master List who've submitted their Health form
  for (var i = 0; i < HealthForms.length; ++i) {
    // Getting person's name and searching for name on Master List
    Person = String(HealthForms[i][0] + " " + HealthForms[i][1]);
    personGrade = HealthForms[i][2];
    personIndex = LookForName(Person, Members);
    
    // Checking off person's name on Master List
    if (personIndex >= 0) {
      grades[personIndex] = [personGrade];
      HealthConfirmed[personIndex] = ['X'];
    }
    else
      ReportError(ErrorForm.getLastRow()+1, "Health Form Check: "+Person+" not found in Master List", Person);
  }

  // Setting 'X's on each person who've submitted their Health forms on Master List
  MasterList.getRange("Master List!C2:C"+(2+Members.length-1)).setValues(grades);
  MasterList.getRange("Master List!E2:E"+(2+Members.length-1)).setValues(HealthConfirmed); 
  
  // Formatting spreadsheet
  ErrorForm.getRange("A2:A").setHorizontalAlignment("left");
  MasterList.getRange("Master List!C2:E").setHorizontalAlignment("center");


  
  /* TODO
  - Update list comprehension to javascript standards     
  
  Other cool features
  - Decorating the docs, formatting
  - Auto-emails to groups for missing information?
  
  Quick Test Variables (2017)
  var MasterList = SpreadsheetApp.openById("1p15iP-b1xHR9w58cr4wKxx_Ke3SM4k-jbkQ_wCEV6ZY");
  var TeamInfo = SpreadsheetApp.openById("1jOrSfL7EtQ-cGbSNl4k-EQZzk7Gzwf8oJP9W2qYBwpA");
  var RulesAndWaivers = SpreadsheetApp.openById("1BD_7eNaigyrPSCWH3q2sl4UsniyTRW727SASKLG3Zrg");
  var Health = SpreadsheetApp.openById("1SD11K2VjASqhuxoS60CO0Yiex0R01BR2-M9gWSCXJ6Q");
  var ErrorForm = SpreadsheetApp.openById("1xQeO7mhg9dnuo7bdOe4WMws9ahRmdvkSIFJAbXeh2kY");
   */
}