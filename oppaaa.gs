/* Author: Hong Yoon Kim 
Operation Program for Prompt And Automated Announcement.
https://trello.com/b/aKGMbc2g/operation-program-for-prompt-and-automated-announcement
*/

// collection of utility functions


function getTimestampCol(){
  return 0
}

function getTypesOfEventCol(){
  return 1
}

function getEventNameCol(){
  return 2
}

function getStartDateCol(){
  return 3
}

function getTimeCol(){
  return 4
}


function getLocationCol(){
  return 5
}

function getDescriptionCol(){
  return 6
}

function getEndDateCol(){
  return 7
}

function getDuplicateEventDateCol(){
  return 8
}

function getAdditionalDateCol(){
  return 9
}

function getIntendedAudienceCol(){
  return 10
}

function parseMultipleDate(Dates){
  return Dates.split(";")
}

function isEmpty(input){
  if (input==''){
    return true;
  } else {
    return false;
  }
}


// copy array of event array to one sheet.
function pasteArrayOfEventArray(ArrayOfEventArray, oneSheet){
  var lastRow = oneSheet.getLastRow();
  var numRow = ArrayOfEventArray.length;
  var numCol = ArrayOfEventArray[0].length;
  
  var currentRange = oneSheet.getRange(lastRow+1, 1, numRow, numCol);
  currentRange.setValues(ArrayOfEventArray);
}

function getUpcomingSundayDateTitle(){
  var today = new Date();
  var day = today.getDay();
  var diff = today.getDate() + (7 - day);
  var upcomingSunday = new Date(today.setDate(diff));
  
  var year = upcomingSunday.getYear();
  var month = parseInt(upcomingSunday.getMonth()+1).toString();
  if (month.length == 1){
    month = '0'+month;
  }
  var date = parseInt(upcomingSunday.getDate()).toString();
  if (date.length == 1){
    date = '0'+date;
  }
    
  var title = year +'.' + month +'.' + date;  
  return title
}


function getArrayOfEventArray(data){
    var ArrayOfEventArray = [];
  // required fields: TypesOfEvents, EventName, StartDate, Description, IntendedAudience
  // optional fields: Time, Location, EndDate, DuplicateEventsDate, AdditionalDate
  
  // required fields 
  var TypesOfEvent = data[getTypesOfEventCol()]
  var EventName = data[getEventNameCol()];
  var StartDate = new Date(data[getStartDateCol()]);
  var Description = data[getDescriptionCol()];
  var IntendedAudience = data[getIntendedAudienceCol()];
  var Status = 'Pending';
  
  // optional fields
  var Time = data[getTimeCol()];
  var Location = data[getLocationCol()];
  
  // see if 'End date' is specified. If yes, get duration as well.
  if (data[getEndDateCol()]){
    var EndDate = new Date(data[getEndDateCol()]);
    var Duration = EndDate.getTime() - StartDate.getTime();
  }
  
  else{
    var EndDate = StartDate;
    var Duration = '';
  }
  var DuplicateEventDate = data[getDuplicateEventDateCol()];
  var AdditionalDate = data[getAdditionalDateCol()];
  
  // create 'EventArray'
  var EventArray = [TypesOfEvent, EventName, StartDate.toLocaleDateString("en-US"), EndDate.toLocaleDateString("en-US"), Time, Location, Description,IntendedAudience, Status];
  ArrayOfEventArray.push(EventArray);
  
  // Duplicate or Generate additional events 
  if (DuplicateEventDate) {
    var ArrayDuplicateEventDate = parseMultipleDate(DuplicateEventDate);
    
    for (i=0;i<ArrayDuplicateEventDate.length;i++){
      var NewStartDate = new Date(ArrayDuplicateEventDate[i]);
      
      if (Duration){
        var NewEndDate = new Date(NewStartDate.getTime() + Duration);
      }
      else {
        var NewEndDate = NewStartDate;
      }
      var EventArray = [TypesOfEvent, EventName, NewStartDate.toLocaleDateString("en-US"), NewEndDate.toLocaleDateString("en-US"), Time, Location, Description, IntendedAudience, Status];
      ArrayOfEventArray.push(EventArray);
    }
  }
  
  // if there exists additonal date entry
  if (AdditionalDate) {
    var ArrayAdditional = parseMultipleDate(AdditionalDate);
    for (i=0;i<ArrayAdditional.length;i++){
      var OneAdditional = ArrayAdditional[i];
      var ArrayOneAdditional = OneAdditional.split(":");
      var title = ArrayOneAdditional[0];
      var NewStartDate = new Date(ArrayOneAdditional[1]);
      var NewEndDate = NewStartDate;
      
      var EventArray = [TypesOfEvent, title, NewStartDate.toLocaleDateString("en-US"), NewEndDate.toLocaleDateString("en-US"), Time, Location, Description, IntendedAudience, Status];
      ArrayOfEventArray.push(EventArray);
    }
  }
  return ArrayOfEventArray
}


// move expired events from 'Current' sheet to 'Archive' sheet. 
function updateCurrArchSheet(){
  var ss = SpreadsheetApp.getActive()
  var CurrentSheet = ss.getSheetByName('Current');
  var ArchiveSheet = ss.getSheetByName('Archive');
  
  var LastRow = CurrentSheet.getLastRow();
  var LastCol = CurrentSheet.getLastColumn();
  var SortRange = CurrentSheet.getRange(2,1,LastRow-1,LastCol);
  
  // sort by end date
  SortRange.sort({column: 4});
  
  var i = 2;
  var NumOfRowsToMove = 0;
  var Now = new Date();
  var EndDate = new Date(CurrentSheet.getRange(i,4).getValue());

  while (EndDate < Now){
    i = i+1;
    EndDate = new Date(CurrentSheet.getRange(i,4).getValue());
    NumOfRowsToMove++ ;
  }
  
  // move to the archive sheet
  if (NumOfRowsToMove > 0){
    var MoveRange = CurrentSheet.getRange(2, 1, NumOfRowsToMove, LastCol);
    var TargetRange = ArchiveSheet.getRange(ArchiveSheet.getLastRow()+1,1,NumOfRowsToMove, LastCol);
    MoveRange.moveTo(TargetRange);
  }
  SortRange.sort({column: 4});
}

function getNweeksFromNow(N){
  var nWeeksFromNow = new Date();
  nWeeksFromNow.setDate(nWeeksFromNow.getDate() + 7*N);
  return nWeeksFromNow;
}

function getCurrentEventArray(){
  var ss = SpreadsheetApp.getActive()
  var CurrentSheet = ss.getSheetByName('Current');
  
  var now = new Date();
  var twoWeeksFromNow = new Date();
  twoWeeksFromNow.setDate(twoWeeksFromNow.getDate() + 14);
  
  lastRow = 2;
  numOfRow = 0;
  var endDate = new Date(CurrentSheet.getRange(lastRow,4,1,1).getValue());
  while (endDate <= twoWeeksFromNow){
    lastRow = lastRow + 1;
    numOfRow = numOfRow + 1;
    endDate = new Date(CurrentSheet.getRange(lastRow,4,1,1).getValue());
  }
  
  var lastCol = CurrentSheet.getLastColumn();
  var currentEventArray = CurrentSheet.getRange(2, 1, numOfRow, lastCol);
  
  return currentEventArray.getValues();
}

function getEventObject(){
  var eventObject = {'whatsNew': [], 'sunday':[], 'access':[], 'weeklyEmail':[], 'tc':[]}
  var currentEventArray = getCurrentEventArray();
  for (var i = 0; i< currentEventArray.length; i++){
    var oneEvent = currentEventArray[i];
    placeOneEventInEventObject(oneEvent,eventObject);
  }
  
  return eventObject
}



function placeOneEventInEventObject(oneEvent, eventObject){
  var type = oneEvent[0];
  var startDate = new Date(oneEvent[2]);
  var audience = oneEvent[7];
  
  //var now = new Date();
  var oneWeekFromNow = getNweeksFromNow(1);
  
  // what's new
  if (type!='regular events' && audience=='All'){
    eventObject.whatsNew.push(oneEvent);
  }
  
  
  // Weekly Event Gathering
  if (startDate.getTime() < oneWeekFromNow){
    if (audience=='All'){
      //Sunday
      eventObject.weeklyEmail.push(oneEvent);
      eventObject.sunday.push(oneEvent);
    } else if (audience=='Campus'){
      //Access
      eventObject.access.push(oneEvent);
    } else if (audience=='TC'){
      //TC
      eventObject.tc.push(oneEvent);
    }
  } else {
    eventObject.tc.push(oneEvent);
  }
  
  
  // what's new;
  //// no regular events 
  //// just title
  
  // Sunday 
  /// LIFE Group 
  
  // TC
  //// event name + date
  //// intended audience + this week / week ++
  
 // non-events
 // regular events
 // special events
 // retreats
 // classes
 // application
 // due dates
  
  // weekly email reminder ; sunday; announcement for tc (this week, week ++)
  // Weekly Email Reminders
  // Sunday (2/28); CHURCH-WIDE MINISTRY and CAMPUS MINISTRY
  // Annc. for TC; This week AND planning ahead
  // All, Comm., Camp,
}


function isTitleInArray(eventArray, title){
  var flag = false;
  for (var i=0;i< eventArray.length;i++){
    var oneEvent = eventArray[i];
    if (oneEvent == title){
      flag = true
    }
  }
  return flag
}

function isAccessInArray(array){
  return isTitleInArray(array,'ACCESS');
}


function formatAMPM(date) {
  var hours = date.getHours();
  var minutes = date.getMinutes();
  var ampm = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12;
  hours = hours ? hours : 12; // the hour '0' should be '12'
  minutes = minutes < 10 ? '0'+minutes : minutes;
  var strTime = hours + ':' + minutes + ' ' + ampm;
  return strTime;
}

function formatMonth(int){
  var num = int+1
  return num.toString()
}

//Global Key Dictionary for GDrive integration

//var scriptProperties = PropertiesService.getScriptProperties();
//Logger.log(scriptProperties.getProperties())
//scriptProperties.setProperty('oppaaId', '0B3oTqpvOcBo5YklicWZySWxJcGM')
//scriptProperties.setProperty('weeklyId', '0B3oTqpvOcBo5WGlILTA5V0xLYWs')


function retrieveId (key){
  var scriptProperties = PropertiesService.getScriptProperties();
  var newKey = key + 'Id';
  return scriptProperties.getProperty(newKey)
}

function updateId(key,value){
  var scriptProperties = PropertiesService.getScriptProperties();
  var newKey = key + 'Id';
  scriptProperties.setProperty(newKey,value);
}

function retrieveUrl (key){
  var scriptProperties = PropertiesService.getScriptProperties();
  var newKey = key + 'Url';
  return scriptProperties.getProperty(newKey)
}

function updateUrl (key, value) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var newKey = key + 'Url';
  scriptProperties.setProperty(newKey,value);
}


//GDrive integration.
// 1. make current week folder
// 2. Weekly Announcement doc; write weekly Announcement doc

function makeCurrentWeekFolder(){
  //var oppaaaId = '0B3oTqpvOcBo5YklicWZySWxJcGM';
  //var oppaaaFolder = DriveApp.getFolderById(oppaaaId);
  //oppaaaFolder.createFolder('weekly-annc-folder');
  var weeklyId = '0B3oTqpvOcBo5WGlILTA5V0xLYWs';
  var weeklyFolder = DriveApp.getFolderById(weeklyId);
  var title = getUpcomingSundayDateTitle();
  var currentWeekFolder = weeklyFolder.createFolder(title);
  Logger.log(currentWeekFolder.getId())
  updateId('currentWeek',currentWeekFolder.getId());
}

//// Weekly Announcement
function makeWeeklyAnnoucementFile(){
  var currentWeekId = retrieveId('currentWeek');
  var currentWeekFolder= DriveApp.getFolderById(currentWeekId);
  
  var rootFolder = DriveApp.getRootFolder();
  var title = getUpcomingSundayDateTitle() + " Weekly Announcement";
  
  var fileId = DocumentApp.create(title).getId();
  updateId("weeklyAnnouncement",fileId);
  
  var file =  DriveApp.getFileById(fileId);
  currentWeekFolder.addFile(file)
  rootFolder.removeFile(file);
  
  var url = file.getUrl();
  updateUrl("weeklyAnnouncement",url);
}


//GDocs Integration
// WriteSection 
// // WriteDefault
// // WriteMultipleEventArray
// // // WriteEventArray
// // // // Get Header 
// // // // Write Header -> getHeader
// // // // Write Multiple Description
// // // // // Write Description
// // // // // // write ListItem

// okay this has been a disaster; formatting and writing in google docs is not as simple as I hope; but gotta do what I need to do.

function getHeader(oneEvent){
  var title = oneEvent[1];
  var startDate = new Date(oneEvent[2]);
  var endDate = new Date(oneEvent[3]);
  if (oneEvent[4]){ 
    var time = oneEvent[4];
  } else {
    var time = null;
  }
  var location = oneEvent[5];
  var description = oneEvent[6];
  var arrayDescription = description.split(';');
  var audience = oneEvent[7];
  
  var days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  if (startDate.getTime() == endDate.getTime()){
  var header = title + " (" + days[startDate.getDay()] + ', '
    +  formatMonth(startDate.getMonth()) + "/" + startDate.getDate().toString() + ")";
    if (time != null){
      if (time instanceof Date){
        header += " - " + formatAMPM(time);
      } else{
        header += " - " + time;
      }
    }
  } else {
  var header = title +" (" + days[startDate.getDay()] + "-" + days[endDate.getDay()] + ', '
    + formatMonth(startDate.getMonth()) + "/" + startDate.getDate().toString() +'-'
    + formatMonth(endDate.getMonth()) + "/" + endDate.getDate().toString() + ")" ;
  }
  
  if (!isEmpty(location)){
    header += " @ " + location;
  }
  return header;
}

function writeListItem(body, text, nestingLevel, glyph){
  var listItem = body.appendListItem(text).setNestingLevel(nestingLevel).setGlyphType(glyph);
  return listItem
}

function writeHeader(body, header, listItems){  
  listItems.header = writeListItem(body,header,0,DocumentApp.GlyphType.BULLET);
}

function writeDescription(body, description, listItems){
  listItems.description = writeListItem(body,description,1,DocumentApp.GlyphType.HOLLOW_BULLET);
}

function writeMultipleDescription(body,multipleDescription, listItems){
  for (var i=0; i<multipleDescription.length;i++){
    writeDescription(body,multipleDescription[i], listItems);
  }
}

function writeSubdescription(body, sub, listItems){
  listItems.subDescription = writeListItem(body,sub,2,DocumentApp.GlyphType.SQUARE_BULLET, listItems);
}

function writeMultipleSubdescription(body, multipleSub, listItems){
  for (var i=0; i<multipleSub.length;i++){
    writeSubdescription(body,multipleSub[i], listItems);
  }
}

// here 'descriptions' is from user input from google form; where as multipleDescription is individual description in an array form.
function writeDescriptions(body, descriptions, listItems){
    var multipleDescription = descriptions.split('; ');
    writeMultipleDescription(body,multipleDescription, listItems);
}

function writeHeaderDescriptions(body, header, descriptions, listItems){
  if (header){
    var headerListItem = writeHeader(body,header, listItems);
  }
  
  if (descriptions){
    var descriptionsListItem = writeDescriptions(body,descriptions, listItems);
  }
}

function writeEventArray(body, oneEvent, listItems){
  var header = getHeader(oneEvent);
  var descriptions = oneEvent[6];
  
  writeHeaderDescriptions(body,header,descriptions,listItems);
}

function writeMultipleEventArray(body,multipleEventArray,listItems){
  for (var i=0; i<multipleEventArray.length;i++){
    var oneEventArray = multipleEventArray[i];
    writeEventArray(body,oneEventArray, listItems);
  }
}


function writeSection(body, intro, default_items, multipleEventArray, header_only){
  if (typeof(header_only)==='undefined') header_only = false;
  
  body.appendParagraph(intro);
  var listItems = {'header':null,'description':null,'subDescription':null};
  if (default_items){
    for (var i=0; i< default_items.length; i++){
      default_items[i](body,listItems);
    }
  }
  // initialize List Item as an object; need to be able to pass it on a cascading method.
  
  if (header_only){
    for (var i=0; i<multipleEventArray.length;i++){
      var header = getHeader(multipleEventArray[i])
      writeHeader(body, header, listItems)
    }  
  } else {
    writeMultipleEventArray(body,multipleEventArray,listItems);
  }
  
  // set GlyphType
  setListItemsGlyphType(listItems);
  body.appendParagraph('').appendHorizontalRule();
}
  

function setListItemsGlyphType(listItems){
  if (listItems.subDescription){
    var listItem = listItems.subDescription;
    listItem.setGlyphType(DocumentApp.GlyphType.SQUARE_BULLET)
  }
  
  if (listItems.description){
    var listItem = listItems.description;
    listItem.setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET);
  }
  
  if (listItems.header){
    var listItem = listItems.header;
    listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
  }
}

function debug(){
  var scriptProperties = PropertiesService.getScriptProperties();
  var doc = DocumentApp.openById(retrieveId("weeklyAnnouncement"));
  var body = doc.getBody();
  var currentEventArray = getCurrentEventArray();
  
  //writeSection(body,'testing writeSection function',currentEventArray,[writeLifeGroupSignUp,writeOpenChapelTimes]);
  writeSection(body,'testing writeSection function',[writeLifeGroupSignUp,writeOpenChapelTimes],currentEventArray);
  //writeLifeGroupSignUp(body);
}

/*
function getEventObject(){
  var eventObject = {'whatsNew': [], 'sunday':[], 'access':[], 'weeklyEmail':[], 'tc':[]}
  var currentEventArray = getCurrentEventArray();
  for (var i = 0; i< currentEventArray.length; i++){
    var oneEvent = currentEventArray[i];
    placeOneEventInEventObject(oneEvent,eventObject);
  }
}
*/

function writeWeeklyAnnouncement(){
  var scriptProperties = PropertiesService.getScriptProperties();
  var doc = DocumentApp.openById(retrieveId("weeklyAnnouncement"));
  var body = doc.getBody()
  var eventObject = getEventObject();
  
  // need to write a special logic for simplified version
  
  writeSection(body, 'Weekly Email Reminders', [writeOpenChapelTimes], eventObject.weeklyEmail)
  writeSection(body, 'Sunday',[writeOpenChapelTimes, writeLifeGroupSignUp], eventObject.sunday)
  // need to write is Access Logic here need more complicated Logic here 
  writeSection(body, 'Access',[], eventObject.access)
  writeSection(body, 'TC', [], eventObject.tc, true)
  
}

function writeLifeGroupSignUp(body,listItems){
  var header = 'LIFE Group Sign-up';
  var multipleDescription = ['Sign up for LIFE Group today!',
                      'If you have a smartphone, please sign up on your mobile device at annarborhmcc.net/lifegroups OR sign up by card',
                      'Questions: Contact annarbor@hmcc.net'];
  writeHeader(body,header,listItems);
  writeMultipleDescription(body,multipleDescription,listItems);
  setListItemsGlyphType(listItems);
}

function writeOpenChapelTimes(body,listItems){
  var header = 'Open Chapel Times @ Transformation Center';
  var multipleDescription = ['Morning Prayer (Mon-Fri, 6:30-7:45am)',
                      'Evening Chapel (Mon-Thu, 9-10pm)'];
  var multipleSubdescription = ['There is no Evening Chapel during the UM Spring Break (2/29 - 3/3). Chapel will resume on Monday 3/7.'];
  
  writeHeader(body,header,listItems);
  writeMultipleDescription(body,multipleDescription,listItems);
  //writeMultipleSubdescription(body,multipleSubdescription,listItems);
  setListItemsGlyphType(listItems);
}

//// Pastor's Pespective 
function makePastorsPerspectiveFile(){
  //var oppaaaId = '0B3oTqpvOcBo5YklicWZySWxJcGM';
  var currentWeekId = retrieveId('currentWeek');
  var currentWeekFolder= DriveApp.getFolderById(currentWeekId);
  
  var rootFolder = DriveApp.getRootFolder();
  var title = getUpcomingSundayDateTitle() + " Pastor's Perspective";
  
  var fileId = DocumentApp.create(title).getId();
  updateId("Pastor's Perspective",fileId);
  
  var file =  DriveApp.getFileById(fileId);
  currentWeekFolder.addFile(file)
  rootFolder.removeFile(file);
  
  var url = file.getUrl();
  updateUrl("Pastor's Perspective",url);
}

function solicitPastorsPerspective(){ 
  var url = retrieveUrl("Pastor's Perspective");
  var doc = DocumentApp.openByUrl(url);
  var body = doc.getBody();
  
  if (body.getText() == '') {
    slackTo('oppaaa',"don't forget to write <" + url + " | Pastor's Perspective>");
  }
}

// Slack integration
function slackTo(channel, msg) {
  var SLACK_URL = retrieveUrl('slack');
  var payload = 'payload={"channel": "#' + channel + '","username": "oppaaa", "text": "' + msg + '","icon_emoji": ":kiss:"}';
  var options = 
      {
        "method":"post",
        "payload":payload
      };
  UrlFetchApp.fetch(SLACK_URL,options);
}

function slackPastorsPerspective(){
  var url = retrieveUrl("Pastor's Perspective");
  Logger.log(url)
  var message = "Pastors, please produce prologue of <" + url + "|Pastor's Perspective>"
  slackTo('oppaaa',message)
}

function slackWeeklyAnnouncement(){
  var url = retrieveUrl("weeklyAnnouncement")
  var message = "OPPAAA! Please check <" + url + "|OPPAAA>"
  slackTo('oppaaa',message)
}


function testOppaaa(){
  //makeCurrentWeekFolder();
  
  //// Pastor's Perspective
  makePastorsPerspectiveFile();
  slackPastorsPerspective();
  
  // weekly announcement
  updateCurrArchSheet();
  makeWeeklyAnnoucementFile();
  writeWeeklyAnnouncement();
  slackWeeklyAnnouncement()
}

///////////////////////////////
//// Event-driven functions////
///////////////////////////////
function onFormSubmit(e) {
  
  // initialize variables
  var newSubmitRange = e.range;
  var ss = SpreadsheetApp.getActive()
  var currentSheet = ss.getSheetByName('Current')
  var data = e.values;
  
  // copy to current sheet
  pasteArrayOfEventArray(getArrayOfEventArray(data), currentSheet);
}



//////////////////////////////
//// Time-driven functions////
//////////////////////////////

function runEveryMonday(){
  // remove old events from Current Sheet to Archive Sheet;
  updateCurrArchSheet();
  
  // make current week folder in Google Drive 
  makeCurrentWeekFolder();
  
  // make Pastor's Perspective 
  makePastorsPerspectiveFile();
  
  // slack Pastor's Perspective
  slackPastorsPerspective(); 
}

function everyWednesday(){
  // solicit PastorsPerspective
  solicitPastorsPerspective();
}

// Function for debugg

function DevFormSubmit(e){
  var data = e.values;
}

function letsTestSlack(){  
  // make one week folder in Google Drive 
  var oneWeekFolder = makeOneWeekFolder();
  
  // make Pastor's Perspective 
  var URL_PP = makePastorsPerspectiveFile(oneWeekFolder);
  
  // slack Pastor's Perspective
  slackPastorsPerspective(URL_PP); 
}