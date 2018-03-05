// Opening Spreadsheets
var MasterList = SpreadsheetApp.openById("1NSXptUb7xLx7dl8M19sUiut3PaeGTsJNlbdlnTdfy9s");
var TeamInfo = SpreadsheetApp.openById("1mLdxyBIYJh3BYg1d2N6wFtsoSltUiSG_Wlh_e2ix2vw");
var RulesAndWaivers = SpreadsheetApp.openById("1TxTCjQ67K6bx2_VWOKVxN4a1VK1ypRDkTWhLi-8wB-g");
var Health = SpreadsheetApp.openById("1ouYTtNzyziYOGykPLRTJaUM76BiteKqxVR8cirPNZFc");
var ErrorForm = SpreadsheetApp.openById("1e3zJ-_UXJuhvjOSomgsnCcYaN1D-_AGVGsbjcBYIfpI");

function Lemon() {
  // Automatically organizes new Team info from jotform onto Master List spreadsheet
  // Getting Team information
  var TeamIndex = TeamInfo.getLastRow(); // Index of new Team info           
  var MasterListIndex = MasterList.getSheetByName("Master List").getLastRow();                  // Index of Master List to input new team
      MasterListIndex += (MasterListIndex == 1 ? 1 : 2);                                        // Adding spacing between teams (excluding first team)
  var CeremonyShirtsIndex = MasterList.getSheetByName("For Opening Ceremony").getLastRow() + 1; // Index of Ceremony/Shirts sheet to input new team
  var Team = TeamInfo.getRange("B"+TeamIndex+":CI"+TeamIndex).getValues()[0];                   // Team info
  
  /***   Master List   ***/
  // Setting Team Name
  MasterList.getRange("A"+MasterListIndex).setValue(Team[0]);
    
  // Sanitizing and splitting Team Members into array
  var TeamMembers = [[n.replace(/[^a-z ]/gi, '')] for each (n in String(Team[6]).split("\n"))]; 
    
  // Setting Team Members
  MasterList.getRange("B"+MasterListIndex+":B" + (MasterListIndex+TeamMembers.length-1)).setValues(TeamMembers);
  
  // Getting Coaches
  var Coaches = [];
  for (var n = 14; n < Team.length; n+=5){ 
    //Ensures empty coaches are not included
    if (Team[n] != "")
      Coaches.push([Team[n] + " " + Team[n+1]]);
  }
  
  // Setting Team Coaches
  MasterList.getRange("E"+(MasterListIndex)+":E"+(MasterListIndex+Coaches.length-1)).setValues(Coaches);
  
  /***   Opening Ceremony   ***/
  MasterList.getRange("For Opening Ceremony!A"+CeremonyShirtsIndex).setValue(Team[0]); // Setting Team Names
  MasterList.getRange("For Opening Ceremony!B"+CeremonyShirtsIndex).setValue(Team[7]); // Setting Team Songs 
  MasterList.getRange("For Opening Ceremony!C"+CeremonyShirtsIndex).setValue(Team[9]); // Setting Team Descriptions 
  
  /***   Shirts   ***/
  // Setting Team Names 
  MasterList.getRange("Shirts!A"+CeremonyShirtsIndex).setValue(Team[0]);
  
  // Cleaning shirt input
  var TeamShirts = [[n.slice(-1) for each (n in String(Team[11]).split("\n"))]];
  
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


function LookForName(name, list) {
  // Searches for name in list (of people already signed up for the relay on Master List)
  for (var n = 0; n < list.length; n++){
    if (name.toUpperCase() == list[n][0].toUpperCase()) {
      return { //Returns true and index of name on Master List
        "present":true,
        "index":n
      };
    }
  }
  return {"present": false};
}


function RulesAndWaiversCheck() {
  // Automatically checks off person who has submitted their Rules and Waivers Form
  // Declaring Variables
  var Members = MasterList.getRange("B2:B"+MasterList.getLastRow()).getValues();                                         // List of people already signed up on relay on Master List
  var Person = RulesAndWaivers.getRange("C"+RulesAndWaivers.getLastRow()+":D"+RulesAndWaivers.getLastRow()).getValues(); // Person who submitted form
      Person = String(Person[0][0] + " " + Person[0][1]);
  var NamePresent = LookForName(Person, Members);                                                                        // Searches for person on Master List
  
  // Checking off person on Master List
  if (NamePresent.present == true)
    MasterList.getRange("Master List!C"+(2+NamePresent.index)).setValue('X');
  else {
    // If person is not found on Master List, error will be reported to be fixed manually later
    var Timestamp = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'"); 
    var ErrorIndex = ErrorForm.getLastRow() + 1;
    var Error = [[Timestamp, "Rules and Waivers Form Check: "+Person+" not found in Master List", Person]];
    ErrorForm.getRange("A"+ErrorIndex+":C"+ErrorIndex).setValues(Error);
  }
}


function HealthCheck() {
  // Automatically checks off person who has submitted their Health Form
  // Declaring Variables
  var Members = MasterList.getRange("B2:B"+MasterList.getLastRow()).getValues();              // List of people already signed up on relay on Master List
  var Person = Health.getRange("C"+Health.getLastRow()+":D"+Health.getLastRow()).getValues(); // Person who submitted form
      Person = String(Person[0][0] + " " + Person[0][1]);
  var NamePresent = LookForName(Person, Members);                                             // Searches for person on Master list
  
  // Checking off person on Master List
  if (NamePresent.present == true)
    MasterList.getRange("Master List!D"+(2+NamePresent.index)).setValue('X');
  else {
    // If person is not found on Master List, error will be reported to be fixed manually later
    var Timestamp = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
    var ErrorIndex = ErrorForm.getLastRow() + 1;
    var Error = [[Timestamp, "Health Form Check: "+Person+" not found in Master List", Person]];
    ErrorForm.getRange("A"+ErrorIndex+":C"+ErrorIndex).setValues(Error);
  }
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
   * Recommendation: Make copies of spreadsheets and save data from each year before clearing for future developers
   */
  
  MasterList.getRange("Master List!A2:G").clear();           // Clears Master List
  MasterList.getRange("For Opening Ceremony!A2:C").clear();  // Clears Opening Ceremony sheet
  MasterList.getRange("Shirts!A2:G").clear();                // Clears Shirts sheet
  ErrorForm.getRange("A2:C").clear();                        // Clears Error Form
  TeamInfo.getRange("A2:CP").clear();
  RulesAndWaivers.getRange("A2:Q").clear();
  Health.getRange("A2:BG").clear();
  
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
   *
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
    MasterList.getRange("E"+(nextLine)+":E"+(nextLine+Coaches[i].length-1)).setValues(Coaches[i]);
    
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
  var Person, NamePresent, Timestamp, ErrorIndex; 
  
  // Checking off each person on Master List who've submitted their Rules and Waivers form
  for (var i = 0; i < RulesAndWaiversForms.length; ++i) {
    // Getting person's name and searching for name on Master List
    Person = String(RulesAndWaiversForms[i][0] + " " + RulesAndWaiversForms[i][1]);
    NamePresent = LookForName(Person, Members);
    
    // Checking off person's name on Master List
    if (NamePresent.present == true)
      RulesAndWaiversConfirmed[NamePresent.index] = ['X'];
    else { // If name not found on Master List, reports the error to be fixed manually later
      Timestamp = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
      ErrorIndex = ErrorForm.getLastRow() + 1;
      Error = [[Timestamp, "Rules and Waivers Form Check: "+Person+" not found in Master List", Person]];
      ErrorForm.getRange("A"+ErrorIndex+":C"+ErrorIndex).setValues(Error);
    }
  }
  
  // Setting 'X's on each person who've submitted their Rules and Waivers forms on Master List
  MasterList.getRange("Master List!C2:C"+(2+Members.length-1)).setValues(RulesAndWaiversConfirmed);
  
  /*** Health Forms ***/
  var HealthForms = Health.getRange("C2:D"+Health.getLastRow()).getValues();   // List of people who've submitted their Health forms
  for(var HealthConfirmed = []; HealthConfirmed.length < Members.length; HealthConfirmed.push([" "])); // Setting up an empty array to track people checked on Master List
  
  // Checking off each person on Master List who've submitted their Health form
  for (var i = 0; i < HealthForms.length; ++i) {
    // Getting person's name and searching for name on Master List
    Person = String(HealthForms[i][0] + " " + HealthForms[i][1]);
    NamePresent = LookForName(Person, Members);
    
    // Checking off person's name on Master List
    if (NamePresent.present == true)
      HealthConfirmed[NamePresent.index] = ['X'];
    else { // If name not found on Master List, reports the error to be fixed manually later
      Timestamp = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
      ErrorIndex = ErrorForm.getLastRow() + 1;
      Error = [[Timestamp, "Health Form Check: "+Person+" not found in Master List", Person]];
      ErrorForm.getRange("A"+ErrorIndex+":C"+ErrorIndex).setValues(Error);
    }
  }

  // Setting 'X's on each person who've submitted their Health forms on Master List
  MasterList.getRange("Master List!D2:D"+(2+Members.length-1)).setValues(HealthConfirmed); 
}
