// September 3, 2013 - Initial Working Version
// AUTHOURS: Akinwale Olaleye, wale@railsfever.com
//           Justin Duffy, justin.duffy@worldlearning.org
//           Dave Taylor, tayloredwebsites@me.com
////////////////////////////////////////////////////////////////////////////////////////////////
// GIVEN:              A spreadsheet with grades, course, semesester, unit, learning outcomes.
// APPLICATION SHOULD: 1) Display a UI for user to filter on above items.
//                     2) Determine the row of the selected learning outcome.
//                     3) Create a lesson plan document based on row data using a template.
//                    
// Main Function is 'lessonPlanUI'. Run using the menu item Lesson Plan -> Start
////////////////////////////////////////////////////////////////////////////////////////////////

 var SOURCE_TEMPLATE = "1Rp-7z7aDcfScLDxmS9pca1gmxDwoX8rUkP8O_RZphN4";

 // In which spreadsheet we have all the course data
 var COURSE_SPREADSHEET = "0AkqXzP2uqsUEdExrUzBtVWNST1FtU0YyUjB4R2xkNnc";

 // In which Google Drive we toss the target documents
 //var TARGET_FOLDER = "0B0qXzP2uqsUETDRZMnJkM2RsYVE";
 var TARGET_FOLDER = "0B0qXzP2uqsUETDRZMnJkM2RsYVE";

 //To support changing column positions
 var GRADE_COLUMN            = 1;
 var COURSE_COLUMN           = 3;
 var SEMESTER_COLUMN         = 2;
 var UNIT_COLUMN             = 4;
 var LEARNING_OUTCOME_COLUMN = 5;
 var COLUMN_NAMES            = 1; //row that contains column names     
 


function _generateLessonPlan(row) {

  var data = _getWorkBook();

  // Fetch variable names, they are column names in the spreadsheet
  var sheet = _getLearningOutcomeSheet(data);
  var columns = getRowAsArray(sheet, COLUMN_NAMES);

  //Logger.log("Processing columns:" + columns);

  var employeeData = getRowAsArray(sheet, row);  
  //Logger.log("Processing data:" + employeeData);

  // Assume first column holds the name of the employee
  var employeeName = employeeData[2];
  var timeStamp = employeeData[0];
  
  var target = createDuplicateDocument(SOURCE_TEMPLATE, employeeName + " Lesson Plan");

  Logger.log("Created new document:" + target.getId());

  for(var i=0; i<columns.length; i++) {
    // TAG forma is :key:
    var key = ":" + columns[i] + ":"; 
    var text = employeeData[i] || ""; // No Javascript undefined
    replaceString(target, key, text);
      
    //var newString = key +" " + text;
    //var newParagraph = key + " " + text;
    //replaceParagraph(target, key, newParagraph);
      
  }

}

 function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [ {name: "Start", functionName: "lessonPlanUI"}];
  ss.addMenu("Lesson Planner", menuEntries);
 }


 
 function lessonPlanUI() { 
   var doc = SpreadsheetApp.getActiveSpreadsheet();
   var app = UiApp.createApplication().setTitle('Course Planner');
   var panel = app.createVerticalPanel();
   panel.setId("main-panel");
   
  
   //grades
   panel.add(_getLabel("Grades"));
   uniqueGrades = _findUniqueGrades();
   var gradeListBox = _getListBox(app,uniqueGrades,"grade-listbox","grades");
   var gradeHandler = app.createServerHandler('_clickGrade').addCallbackElement(panel);
   gradeListBox.addChangeHandler(gradeHandler);
   panel.add(gradeListBox);
   
   //courses
   panel.add(_getLabel("Courses"));
   var courseListBox = _getListBox(app,{},"course-listbox","courses");
   var courseHandler = app.createServerHandler('_clickCourse').addCallbackElement(panel);
   courseListBox.addChangeHandler(courseHandler);
   panel.add(courseListBox);
    
   //semester
   panel.add(_getLabel("Semesters"));
   var semesterListBox = _getListBox(app,{},"semester-listbox","semesters");
   var semesterHandler = app.createServerHandler('_clickSemester').addCallbackElement(panel);
   semesterListBox.addChangeHandler(semesterHandler);
   panel.add(semesterListBox);
     
   //unit
   panel.add(_getLabel("Units"));
   var unitListBox = _getListBox(app,{},"unit-id","units");
   var unitHandler = app.createServerHandler('_clickUnit').addCallbackElement(panel);
   unitListBox.addChangeHandler(unitHandler);
   panel.add(unitListBox);
  
   //learning outcomes
   panel.add(_getLabel("Learning Outcomes"));
   var loListBox = _getListBox(app,{},"lo-listbox","learningOutcomes");
   loListBox.setSize("100%","7%");
   panel.add(loListBox);
   
   //create lesson plan
   var clpButton = app.createButton("Create Lesson Plan");
   var clpButtonHandler = app.createServerHandler('_clickClpButton').addCallbackElement(panel);
   clpButton.addClickHandler(clpButtonHandler);
   panel.add(clpButton);
   
   //hidden fields to track listbox state
   var gradeTracker = app.createHidden().setId("selectedGrade").setName("selectedGrade");
   panel.add(gradeTracker);
   var courseTracker = app.createHidden().setId("selectedCourse").setName("selectedCourse");
   panel.add(courseTracker);
   var semesterTracker = app.createHidden().setId("selectedSemester").setName("selectedSemester");
   panel.add(semesterTracker); 
   var unitTracker = app.createHidden().setId("selectedUnit").setName("selectedUnit");
   panel.add(unitTracker); 
   
   //finally
   app.add(panel);
   doc.show(app);  
 }

 
 function _clickGrade(eventInfo) {
   var app = UiApp.getActiveApplication();
   // get value of ListBox based on name
   var selectedGrade = eventInfo.parameter.grades;
   app.getElementById("selectedGrade").setValue(selectedGrade);
   //Logger.log("eventInfo.parameter.selectedGrade: "+eventInfo.parameter.selectedGrade);
    _displayCourses(selectedGrade);
   return app;
 }

 function _clickCourse(eventInfo) {
   var app    = UiApp.getActiveApplication();
   var grade  = eventInfo.parameter.selectedGrade;
   var course = eventInfo.parameter.courses;
   app.getElementById("selectedCourse").setValue(course);
   //Logger.log("Selected grade, course: "+grade+","+course);
   _displaySemesters(grade,course);
   return app;
 }

 function _clickSemester(eventInfo) {
   var app    = UiApp.getActiveApplication();
   var grade  = eventInfo.parameter.selectedGrade;
   var course = eventInfo.parameter.selectedCourse;
   var semester = eventInfo.parameter.semesters;
   app.getElementById("selectedSemester").setValue(semester);
   //Logger.log("Selected grade, course, semester: " +grade+", "+course+" ,"+semester);
   _displayUnits(grade,course,semester);
   return app;
 }

 function _clickUnit(eventInfo) {
   var app    = UiApp.getActiveApplication();
   var grade  = eventInfo.parameter.selectedGrade;
   var course = eventInfo.parameter.selectedCourse;
   var semester = eventInfo.parameter.selectedSemester;
   var unit = eventInfo.parameter.units;
   //Logger.log("grade, course, semester, unit: "+grade+", "+course+", "+semester+", "+unit);
   _displayLearningOutcomes(grade, course, semester, unit);
   return app;
 } 

 function _clickClpButton(eventInfo){
   var app = UiApp.getActiveApplication();
   var row;
   //get text of selected learning outcome
   selectedOutcome = eventInfo.parameter.learningOutcomes;
   //Logger.log("Selected Outcome :"+selectedOutcome);
   
   //find the row in the sheet that matches the learning outcome
   var sheet = _getLearningOutcomeSheet(_getWorkBook());
   var data  = sheet.getDataRange().getValues(); // read all data in the sheet
   var col   = LEARNING_OUTCOME_COLUMN - 1;
   
   for(var n=0;n<data.length;++n){ 
     if( data[n][col].toString().match(selectedOutcome) == selectedOutcome ) { 
       row = n + 1;
       //Logger.log("matched row: "+ row); 
     }
   }
   
   _generateLessonPlan(row)
   return app;
  }  

 function _displayCourses(grade) {
   var app = UiApp.getActiveApplication();
   var validCourses = _findValidCourses(grade);
   var lb = app.getElementById("course-listbox");
   _populateListBox(lb,validCourses);
 }

 function _displaySemesters(grade,course){
   var app = UiApp.getActiveApplication();
   var validSemesters = _findValidSemesters(grade,course);
   var lb = app.getElementById("semester-listbox");
   _populateListBox(lb,validSemesters);
 }
 
 function _displayUnits(grade,course,semester){
   var app = UiApp.getActiveApplication();
   var validUnits = _findValidUnits(grade,course,semester);
   var mlb = app.getElementById("unit-id");
   //Logger.log("Listbox Widget: "+mlb.getId() );
   _populateListBox(mlb,validUnits);
 }

 function _displayLearningOutcomes(grade, course, semester, unit){
   var app = UiApp.getActiveApplication();
   var validLearningOutcomes = _findValidLearningOutcomes(grade, course, semester, unit);
   var lb = app.getElementById("lo-listbox");
   _populateListBox(lb,validLearningOutcomes);
 }

 function _findUniqueGrades(){
   var graderows = _getGradeRows(); 
   var grades    = {a:'Select a grade'};
   
   for(var i=1; i<graderows.length; i++) { //skip the first row (labels)
     if(graderows[i] != "" && graderows[i] != undefined ) {
       grades[graderows[i]]=graderows[i];
     }
   }
   return grades;
 }

 function _findValidCourses(grade) {
  var graderows  = _getGradeRows();
  var courserows = _getCourseRows();
   
  // loop over courserows
  //  if the grade for this courserow matches the grade we are searching for
  //     then store this courserow in an associative array
  validCourses = {a:'Select a Course'};
  for(var i=0; i< courserows.length; i++) {
    if ( graderows[i] == grade ) {
      validCourses[courserows[i]] = courserows[i];
    }
  }
  return validCourses;
 } 
 
 function _findValidSemesters(grade,course){
   var graderows    = _getGradeRows();
   var courserows   = _getCourseRows();
   var semesterrows = _getSemesterRows(); 
   validSemesters = {a:'Select a Semester'};
   
   for(var i=0; i< semesterrows.length; i++){
     if ( courserows[i] == course && graderows[i] == grade ){
       validSemesters[semesterrows[i]] = semesterrows[i];
     }
   }
   return validSemesters;
 }

 function _findValidUnits(grade,course,semester){
   var graderows    = _getGradeRows();
   var courserows   = _getCourseRows();
   var semesterrows = _getSemesterRows(); 
   var unitrows     = _getUnitRows(); 
   var validUnits = {a:'Select a Unit'};
   
   for(var i=0; i < unitrows.length; i++){
     if (graderows[i] == grade && courserows[i] == course && semesterrows[i] == semester){
       validUnits[unitrows[i]] = unitrows[i];
       //Logger.log("valid unit "+i+" :"+unitrows[i]);
     }
   }
   return validUnits; 
 }
 
 function _findValidLearningOutcomes(grade, course, semester, unit){
   var graderows     = _getGradeRows();
   var courserows    = _getCourseRows();
   var semesterrows  = _getSemesterRows(); 
   var unitrows      = _getUnitRows();
   var lorows        = _getOutcomeRows();
   var validOutcomes = {a:'Select a Learning Outcome'};
   
   for(var i=0;i<lorows.length; i++){
     if(unitrows[i] == unit && graderows[i] == grade && courserows[i] == course && semesterrows[i] == semester ){
        validOutcomes[lorows[i]] = lorows[i];  
     }
   }
   return validOutcomes;
 }

 function _getListBox(app,items,id,label){
   var lb = app.createListBox().setId(id).setName(label);
   lb.setVisibleItemCount(1);
   //add items, if any
   for each (var i in items) {
     lb.addItem(i);
   }
   return lb;
 } 

 function _populateListBox(listBox,items){
   listBox.clear();
   //Logger.log("ListBox ID: "+ listBox.getId());
   for each (var i in items) {
     //Logger.log("adding item: " + i);
     listBox.addItem(i.toString());
   }
 }

 function _getWorkBook(){
   return SpreadsheetApp.openById(COURSE_SPREADSHEET);
 }

 function _getLearningOutcomeSheet(workBook){
   return workBook.getSheets()[0];
 }
 
 function _getLabel(string){
   var app = UiApp.getActiveApplication();
   return app.createLabel(string); 
 }

 function _getGradeRows(){
   var workBook      = _getWorkBook();
   var sheet         = _getLearningOutcomeSheet(workBook); 
   return _getColumnAsArray(sheet, GRADE_COLUMN);
 }

 function _getCourseRows(){
   var workBook      = _getWorkBook();
   var sheet         = _getLearningOutcomeSheet(workBook); 
   return _getColumnAsArray(sheet, COURSE_COLUMN);
 }

 function _getSemesterRows(){
   var workBook      = _getWorkBook();
   var sheet         = _getLearningOutcomeSheet(workBook); 
   return _getColumnAsArray(sheet, SEMESTER_COLUMN);
 }

 function _getUnitRows(){
   var workBook      = _getWorkBook();
   var sheet         = _getLearningOutcomeSheet(workBook); 
   return _getColumnAsArray(sheet, UNIT_COLUMN);
 }

 function _getOutcomeRows(){
   var workBook      = _getWorkBook();
   var sheet         = _getLearningOutcomeSheet(workBook); 
   return _getColumnAsArray(sheet, LEARNING_OUTCOME_COLUMN);
 }

 function _getColumnAsArray(sheet,column) {
  var dataRange = sheet.getRange(1,column,sheet.getLastRow(),1);
  var data = dataRange.getValues()
  //Logger.log("column values: " + data);
  
  return data;
}

/**
 * Return spreadsheet row content as JS array.
 *
 * Note: We assume the row ends when we encounter
 * the first empty cell. This might not be 
 * sometimes the desired behavior.
 *
 * Rows start at 1, not zero based!!!  
 *
 */
function getCurrentRow() {
  var currentRow = SpreadsheetApp.getActiveSheet().getActiveSelection().getRowIndex();
  return currentRow;
}


function getRowAsArray(sheet, row) {
  var dataRange = sheet.getRange(row, 1, 1, sheet.getLastColumn());
  var data = dataRange.getValues();
  var columns = [];

  for (i in data) {
    var row = data[i];

    //Logger.log("Got row: " + row);

    for(var l=0; l<99; l++) {
        var col = row[l];
        // First empty column interrupts
        if(!col) {
            break;
        }

        columns.push(col);
    }
  }

  return columns;
}



/**
 * Duplicates a Google Apps doc
 *
 * @return a new document with a given name from the orignal
 */
function createDuplicateDocument(sourceId, name) {
    var source = DocsList.getFileById(sourceId);
    var newFile = source.makeCopy(name);

    var targetFolder = DocsList.getFolderById(TARGET_FOLDER);
    newFile.addToFolder(targetFolder);

    return DocumentApp.openById(newFile.getId());
}
/**
 * Search a paragraph in the document and replaces it with the generated text 
 */
function replaceParagraph(doc, keyword, newText) {
  var ps = doc.getParagraphs();
  for(var i=0; i<ps.length; i++) {
    var p = ps[i];
    var text = p.getText();

    if(text.indexOf(keyword) >= 0) {
      p.setText(newText);
      p.setBold(false);
      
    }
  } 
}

/**
 * Search a String in the document and replaces it with the generated newString, and sets it Bold
 */
function replaceString(doc, String, newString) {

  var ps = doc.getParagraphs();
  for(var i=0; i<ps.length; i++) {
    var p = ps[i];
    var text = p.getText();
    //var text = p.editAsText();

    if(text.indexOf(String) >= 0) {
      //look if the String is present in the current paragraph
      

      //p.editAsText().setFontFamily(b, c, DocumentApp.FontFamily.COMIC_SANS_MS);
      p.editAsText().replaceText(String, newString);
      
      
      // we calculte the length of the string to modify, making sure that is trated like a string and not another ind of object.
      var newStringLength = newString.toString().length;
      
      // if a string has been replaced with a NON empty space, it sets the new string to Bold, 
      if (newStringLength > 0) {
        // re-populate the text variable with the updated content of the paragraph
        text = p.getText();
        p.editAsText().setBold(text.indexOf(newString), text.indexOf(newString) + newStringLength - 1, true);
      }
    }
  } 
}


 
