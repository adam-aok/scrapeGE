main();


//basic declaration of interaction levels to override link/font warnings to ease batch run
function main(){
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
    //app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
    if(app.documents.length != 0){
		if (app.activeDocument.stories.length != 0){
              myFolder = ("C:/Users/a.keefe/Documents/")
              expFormat = ".txt"          
              myExportPages(expFormat, myFolder)
              $.gc()
		}
    
		else{
			alert("The document does not contain any text. Please open a document containing text and try again.");
		}
	}
	else{
		alert("No documents are open. Please open a document and try again.");
	}
}

function myExportPages(myExportFormat, myFolder){
    
    var curDate = new Date();
    myFileName = "ak.DataMine_" + curDate.toDateString()+ myExportFormat;
    myFilePath = myFolder + "/" + myFileName;
    var myFile = new File(myFilePath);
    
//resolves the lack of indexOf function   
if (typeof Array.prototype.indexOf != "function") {  
    Array.prototype.indexOf = function (el) {  
        for(var i = 0; i < this.length; i++) if(el === this[i]) return i;  
        return -1;  
        }  
}                  
    //var basicMasters = ["A-Master","B.1-Project Team (Headshots)","B.2-Project Team (Group Photo)","B.3-Project Details","C.1-P
      
    var pageRow = [];
    
    //fill row with null values at first
    for(pR = 0; pR < 10 ; pR++) pageRow[pR] = ("\"" + "\"");
        
            
    //input filepath into pageRow
    
    
    //q4 2019 locations array
    var locArr1 = [1.638,2.3902,3.1502,3.8978,4.651,5.4052,6.1586,6.9531]

    
    //file information 
    var fileArr = ["\"" + "\"","\"" + "\"","\"" + "\""]
    
      
    //pulls the first master spread to determine the prototypical layout of the open doc    
    //mSpread = app.activeDocument.masterSpreads.item(0);
    
    //checks for footer title for textbox location options 1 or 3. if match, sets reference array equal
//~     for(a=0;a<mSpread.textFrames.length;a++){
//~         if (mSpread.textFrames.item(a).parentStory.contents == "DESIGN ON THE BOARDS      Q-4 2019") posList= locArr1;
//~         if (mSpread.textFrames.item(a).parentStory.contents == "DESIGN ON THE BOARDS      Q-4 2017") posList = locArr3;
//~         }
        
    //creates null array of projects on the page, to be either modified or added to with pageRow
    //var projIndex = [[]];
    //projIndex.pop()
    
    //load file info array into project index of filepath, modify date
    fileArr[0] = csvQuotes(app.activeDocument.fullName.toString().replace("/g/","G:/"));
    fileArr[1] = csvQuotes(File(app.activeDocument.fullName).modified);
    app.scriptPreferences.measurementUnit = MeasurementUnits.INCHES;
    
         
        
    //iterating through spreads of document (spreads are designated instead of pages to accommodate spreads of multiple sheets with common information)
    for(myCounter = 0; myCounter < app.activeDocument.pages.length; myCounter++){
         
         //get current page
         myPage = app.activeDocument.pages.item(myCounter);         

        
        //checking the master sheet of the current page, assigns it to checkable variable. if not applied, check sheet is converted to masterCheck
        //if (myPage.appliedMaster !== null) masterCheck = myPage.appliedMaster;
        //else masterCheck = app.activeDocument.masterSpreads.item(0);
        
        //modify boolean initialized to assume this page contains a new project
         //modify = false;
         
         //pageRow is initialized as empty, with project values set to empty strings
         pageRow = [];
         for(pR = 0; pR < 5 ; pR++) pageRow[pR] = ("\"" + "\"");
         
         //initialization of margin variable to be used in checking if the layout has been moved from master layout; initializes largest story on page is null, assumes the mastersheet has not been modified and thus need not be checked.
                  
         //ungroups any layers within the document, if there are any.                  
         for(z=0;z<myPage.groups.length;z++){
                checkGroup = myPage.groups.item(z);
                checkGroup.ungroup(); 
            }
        

                         
                                
         for(myCount = 0; myCount < myPage.textFrames.length; myCount++){

                //get textFrame on page, pulls parent story and respective text
                 myTextFrame = myPage.textFrames.item(myCount);
                 //alert(myTextFrame.rotationAngle); --checks for any textframes with irregular orientation
                 myStory = myTextFrame.parentStory;                          
                 myStoryText = csvFriendly(myStory.contents);
                 //alert(myStoryText);
                 
                //initializes paragraph style for that textframe, and sets the style to that of the first paragraph
                 //var myStoryStyle = "undefined";
                 //if (myStory.paragraphs.firstItem().isValid == true) myStoryStyle = myStory.paragraphs.firstItem().appliedParagraphStyle.name;
                 
                 //get coordinates of this text frame for comparison
                 myPosition = myTextFrame.geometricBounds;
                
                    if (approx(myPosition[1],1.006) == true ) pageRow[0] = csvQuotes(myStoryText);
                    if (approx(myPosition[1],3.2444) == true && pageRow[1] == "\"" + "\"") pageRow[1] = csvQuotes(myStoryText);
                    if (approx(myPosition[1],5.4325) == true && pageRow[2] == "\"" + "\"") pageRow[2] = csvQuotes(myStoryText);                        
                    if (approx(myPosition[1],7.9776) == true && pageRow[3] == "\"" + "\"") pageRow[3] = csvQuotes(myStoryText);
                    if (approx(myPosition[1],11.6699) == true && pageRow[4] == "\"" + "\"") pageRow[4] = csvQuotes(myStoryText);
                    }
           
    
  
             
       //debugging code from working out the results 
//~     if (titlesIndex.indexOf(pageRow[0]) == -1){
//~                      if (pageRow[0] !== "\"" + "\"" && projIndex.indexOf(pageRow) == -1){
//~                          projIndex.push(pageRow);
//~                          alert(pageRow);
//~                          }    
//~                      }
//~ ink adding functions seems to have created an issue
//~ var linkArr = [];
//~ parses for link list
//~ for(myCounter = 0; myCounter < app.activeDocument.links.length; myCounter++){
//~     eachLink = app.activeDocument.links.item(myCounter).filePath;
//~     if (linkArr.indexOf("\"" + eachLink + "\"") == -1){
//~         linkArr.push("\"" + eachLink + "\"")
//~     }
        fileArr[2] = myCounter+1;
        myPageText = fileArr.toString() +"\," + pageRow.toString() + "\n";
        //alert(myPageText);
        //alert(titleRow.toString());
        writeFile(myFile,myPageText);
//~     for (j = 0; j< pageRow.length;j++){
//~         //alert(pageRow.toString())
//~         myPageText = fileArr.toString() +"\," + pageRow[j].toString() + "\n";
//~         //alert(myPageText);
//~         //alert(titleRow.toString());
//~         writeFile(myFile,myPageText);
//~         }
}
}

    
  
//sets encoding properties and writing components
function writeFile(fileObj, fileContent, encoding) {  
    encoding = encoding || "UTF-8";
    var titleRow = [csvQuotes("Path"),csvQuotes("Modified"),csvQuotes("Page Number in Document"),csvQuotes("Name"),csvQuotes("Office"),csvQuotes("Project"),csvQuotes("Software"),csvQuotes("Methodology")];
    if (!fileObj.exists) fileContent2 = titleRow.toString() + "\n" + fileContent;
    else fileContent2 = fileContent;
         
    fileObj = (fileObj instanceof File) ? fileObj : new File(fileObj);  
  
    var parentFolder = fileObj.parent;
    if (!parentFolder.exists && !parentFolder.create())  
        throw new Error("Cannot create file in path " + fileObj.fsName);  
  
    fileObj.encoding = encoding;  
    fileObj.open("a");  
    fileObj.write(fileContent2);  
    fileObj.close();
    
    return fileObj;  
}  


//convert text into compatible csv format--REPLACE </P><P> WITH PARAGRAPH BREAK FOR HTML FORMATTING AND #% WITH COMMA
 
 //creates function to compare numbers to see if two are approximately the same
function approx(number,reference,delta){
    var delta = .02;
    if (Math.abs((number - reference)) <= delta || Math.abs((reference - number)) <= delta){
    return true
    }
    else return false
    }

//3 functions to rework text into quotes/HTML format for the purposes of being loaded into a workable CSV
function csvQuotes(myText){
    myText = ("\"" + trim(myText) + "\"");
    return myText
}
function csvFriendly(myText){
    myText = trim(myText.toString().replace(/(\r\n|\n|\r)/gm,"</P><P>").replace(/,/g,"#%"));
    return myText;
 }
function trim(str) {
    return str.toString().replace(/^\s\s*/, '').replace(/\s\s*$/, '');
}

//logging function used for debugging purposes
function logMe(input){
     var now = new Date();
     var output = now.toTimeString() + ": " + input;
     $.writeln(output);
     var logFile = File("/path/to/logfile.txt");
     logFile.open("e");
     logFile.writeln(output);
     logFile.close();
}


