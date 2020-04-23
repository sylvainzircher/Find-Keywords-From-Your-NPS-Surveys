/**
 * Script built with LoVe by Sylvain Zircher - 2019
 */

function onOpen() {
  /**
   * Function to add a new menu to the ui so we can ran the code directly from the Google Sheet.
   */
  
  ui = SpreadsheetApp.getUi();
  
  ui.createMenu("Analyze Comments")
  .addItem("Find keywords...", "FindWords") 
  .addItem("Run stats...", "runStats") 
  .addToUi();
};

function WhatRow(columnName){
   /**
   *Function that finds the last row containing data for a specific column.
   *@param columnName {string}
   *@return the number of the last row containing data.
   *@customfunction
   */
  
  var ss = SpreadsheetApp.getActive();
  var s = ss.getActiveSheet();
  var lc = ss.getLastColumn();
  var lr = ss.getLastRow();
  
  for (var i = 1; i <= lc; i++){
    
    if (s.getRange(1, i).getValue() == columnName) {
    
      data = s.getRange(1,i,lr,1).getValues();
      
      for (var j = data.length - 1; j >= 0; j--) {
        if (data[j][0] != null && data[j][0] != "") {
          return(j + 1);  
        }  
      }
    }
  }
  return undefined;
};

function setColumns() {
 /****************************************************************************************************************
  * HERE WE DEFINE WHICH COLUMN CONTAINS THE COMMENTS (colText). SAME IF YOU HAVE NPS SCORES (colScore)
  * AMEND ACCORDING TO YOUR NEEDS. 
  * COLUMN A = 1
  * COLUMN B = 2
  * COLUMN C = 3  
  * ...
  */
  var colText = 1;
  var colScore = 2;
 /*****************************************************************************************************************/
  
  return([colText, colScore]);
};

function FindWords() {
   /**
   * Function that does not return anything nor take any parameter. From the list of comments and their NPS, it finds all the words
   * that are used, the number of occurences and remove any "Junk" from the list. It then display the data in a nice fashion: ordered
   * with gradient color visual based on the number of occurences. Finally it also calculates the average score associated to the comments
   * where the words appear (only top30).
   * @customfunction
   */  
  var wordsArray = [];
  var wordsCount = [];
  var ss = SpreadsheetApp.getActive();
  var s = ss.getSheetByName("Data"); 
  var lr = s.getLastRow();
  var lc = s.getLastColumn();
  
  var cols = setColumns();
  var colText = cols[0];
  var colScore = cols[1];
  
  var comments = s.getRange(2,colText,lr,1).getValues();
  var i = 0;
  
  comments.forEach(function(text) {
    var words = text.toString().toLowerCase().split(" ");
    words.forEach(function(w) {
      w = w.toString().replace(/\./g,"").replace(/\,/g,"").replace(/\;/g,"").replace(/\!/g,"");
      if (CheckNotJunk(w) && w != "") {
        if(i == 0) {
          wordsArray.push(w);
          wordsCount.push(1);
        } else {
          if(wordsArray.indexOf(w) == -1) {
            wordsArray.push(w);
            wordsCount.push(1);          
          } else {
            var loc = wordsArray.indexOf(w);
            var c = wordsCount[loc];
            wordsCount[loc] = c + 1;
          }
        }
      }
      i = i + 1;      
    });
  });
  
  var waL = wordsArray.length;
  var waC = wordsCount.length;  
  var result = FormatArray(wordsArray, wordsCount);
  var summarySheet = SpreadsheetApp.getActive().getSheetByName("Summary");
  summarySheet.setHiddenGridlines(true);
  
  summarySheet.getRange("A:G").clear();  
  summarySheet.getRange(1,1).setValue("Word");
  summarySheet.getRange(1,2).setValue("Count of appearances");  
  
  if (waL == waC) {
    summarySheet.getRange(2,1,waL,1).setValues(result[0]);
    summarySheet.getRange(2,2,waC,1).setValues(result[1]);
    summarySheet.getRange(1,1,waL,2).sort({column: 2, ascending: false});
     
    // -------- Color Formatting for word frequency 
    var lastRow = waL + 1;
    var rg = "B2:B" + lastRow;
    Logger.log(rg);
    var conditionalFormatRules = summarySheet.getConditionalFormatRules();

    conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
                                .setRanges([summarySheet.getRange(rg)])
                                .whenCellNotEmpty()
                                .setGradientMinpoint('#FFFFFF')
                                .setGradientMaxpoint('#57BB8A')
                                .build());
    
    summarySheet.setConditionalFormatRules(conditionalFormatRules);
  }
};
  
function runStats() {
  /*************************************************************************************
  * HERE WE DEFINE HOW MANY WORDS WE WILL CALCULATE THE AVG NPS SCORE FOR. 
  * BY DEFAULT IT IS SET TO 30 WORDS BUT YOU CAN AMEND THE NUMBER TO WHATEVER YOU WANT.
  * IF YOU UP THE LIMIT IT MIGHT TAKE THE SCRIPT A LOT LONGER TO RUN - ALSO PLEASE BE 
  * AWARE THAT GOOGLE WILL TIME OUT THE REQUEST IF THE QUERY TAKES TOO LONG TO RUN.
  */
  var limit = 20;
  /************************************************************************************/ 
  
  var cols = setColumns();
  var colText = cols[0];
  var colScore = cols[1];
  
  var ss = SpreadsheetApp.getActive();
  var summarySheet = ss.getSheetByName("Summary");
  
  var conditionalFormatRules = summarySheet.getConditionalFormatRules();
  
  StatsPerWord(limit, colText, colScore);
  
  if (summarySheet.getRange(2,7).getValue() != "") {
    summarySheet.getRange(2,WhatRow("Avg Score")).activate();
    conditionalFormatRules.splice(1, 1); /** Remove the first color formatting that was applied **/
    summarySheet.setConditionalFormatRules(conditionalFormatRules);
  }
  
  Logger.log(WhatRow("Nb of comments the Word is in"));
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
                              .setRanges([summarySheet.getRange(2,7,WhatRow("Nb of comments the Word is in") - 1,1)])
                              .whenCellNotEmpty()
                              .setGradientMinpoint('#E67C73')
                              .setGradientMidpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.PERCENTILE, '50')
                              .setGradientMaxpoint('#57BB8A')
                              .build());
  
  summarySheet.setConditionalFormatRules(conditionalFormatRules);
  summarySheet.getRange("A1").activate();
 
};


function FindMatch(text, words) {
  /**
   * Function that checks if the words passed as paramater is included in text.
   * It splits text into all the components it contains and check the words against them. It returns "Yes" if it finds a match, "No" otherwise.
   * Returns "Yes" or "No"
   * @params text (string), words (array)
   * @returns (string) "Yes" or "No"
   * @customfunction
   */  
  var text = text.toString().toLowerCase().split(" ");
  var answer = "";
  
  words.forEach(function(wds){
    wds = wds.toString().toLowerCase();
    text.forEach(function(t) {
      t = t.replace(/\./g,"").replace(/\,/g,"").replace(/\;/g,"").replace(/\!/g,"");   
      if(t == wds){      
        answer = "Yes";    
      }
    });
  });
  
  if (answer != "Yes") {answer = "No";} 
  return answer;
};


function CheckNotJunk(w) {
  /**
   * Function that checks if the word that was passed as a parameter (w) is contained in a list of words.
   * Returns "pass" if it is contained in the list, "go" otherwise.
   * @params w (string)   
   * @returns (string) "pass" or "go"
   * @customfunction
   */
  var list = ["to","and","i", "you", "am", "is", "it", "or", "have", "had", "been", "has", 
              "any", "with", "this", "that", "haven't", "have not", "has not", "as", "the", 
              "a", "it", "it's", "can", "be", "could", "would", "me", "not", "are", "in",
              "on", "if", "but", "for", "of", "but", "all", "my", "from", "was", "were",
              "where", "when", "up", "your", "there", "bit", "will", "like", "very", "lot",
              "load", "at", "do", "does", "don't", "also", "we", "so", "some", "doesn't",
              "out", "too", "an", "by", "i'm", "into", "than", "then", "-", "they", "them",
              "should", "about"];
  if (list.indexOf(w) != -1) { return false;} else {return true;}
};


function FormatArray(w, c) {
   /**
   * Function that returns two arrays of string (respectively the list of words and the list of count) from two arrays of arrays.
   * Returns an array of words (as String) and array of count of words ( as String).
   * @params w (array of arrays) and c (array of arrays).   
   * @returns wordsArrayToReturn (array of String) and wordsCountToReturn (array of String)
   * @customfunction
   */
  var wordsArrayToReturn = [];
  var wordsCountToReturn = [];
  
  w.forEach(function(word) {
    wordsArrayToReturn.push([word]);
  });
  
  c.forEach(function(count) {
    wordsCountToReturn.push([count]);
  });
  
  return [wordsArrayToReturn, wordsCountToReturn];
  
};


function ComputeStats(w, colText, colScore) {
  /**
   * Function that finds out all the comments where the word w passed as parameter is mentioned. Grab the related NPS scores and
   * calculate the average score. Also reports on the numbers of detractor, passive and promoters.
   * @params w (string), colText (integer), colScore (integer)
   * @returns (array) the average nps score for all comments that include the word passed as parameter, number of detractors, passive, 
   * promoters and total number of comments.
   * @customfunction
   */ 
  var ss = SpreadsheetApp.getActive();
  var dataSheet = ss.getSheetByName("Data");
  var lr = dataSheet.getLastRow();
  var colText = colText;
  var colScore = colScore;
  var score = 0;
  var count = 0;
  var detractor = 0;
  var passive = 0;
  var promoter = 0;
  var avgScore = 0;
  var total = 0;
  var text = dataSheet.getRange(2, colText, lr - 1, 1).getValues();
  var scores = dataSheet.getRange(2, colScore, lr - 1, 1).getValues();
  w = [w];
  
  for (var i = 0; i <= lr - 2; i++) {
    if (FindMatch(text[i],w) == "Yes") {
      if (scores[i][0] == "") { /** Handles the case where there is no score data **/
        total = total + 1;
      } else {
        var s = Number(scores[i]);
        score = score + s;
        count = count + 1;
        if (s <= 6) {
          detractor = detractor + 1;
        } else if (s >= 9) {
          promoter = promoter + 1;
        } else {
          passive = passive + 1;
        }
      }
    }
  }
  
  if (count != 0) { /** If scores were provided **/
    avgScore = score / count;
    total = detractor + passive + promoter;
  }
  return [avgScore, detractor, passive, promoter, total];
};


function StatsPerWord(limit, colText, colScore) {
  /**
   * Function that inputs all the statistics for the top n (defined by limit passed as parameter) words in the Summary Sheet. 
   * @params limit (integer), colText (integer), colScore (integer)
   * @customfunction
   */   
  var ss = SpreadsheetApp.getActive();
  var dataSheet = ss.getSheetByName("Data");
  var summarySheet = ss.getSheetByName("Summary");
  var colNb = 10;
  var avgScore = [];
  var detractor = [];  
  var passive = [];
  var promoter = []; 
  var total = [];
  
  if (WhatRow("Nb of comments the Word is in") == undefined) {
    summarySheet.getRange(1,3).setValue("Nb of comments the Word is in");
    summarySheet.getRange(1,4).setValue("Count of Detractors");
    summarySheet.getRange(1,5).setValue("Count of Passive");
    summarySheet.getRange(1,6).setValue("Count of Promoters");
    summarySheet.getRange(1,7).setValue("Avg Score");
    var startRow = 2;
    var offsetRow = limit;
   } else {
     var startRow = WhatRow("Nb of comments the Word is in") + 1;
     var offsetRow = limit;   
   } 
  
  var wds = summarySheet.getRange(startRow,1,offsetRow,1).getValues();
  var result = "";
  
  wds.forEach(function(w, i) {
    result = ComputeStats(w[0], colText, colScore);
    avgScore.push([result[0]]);
    detractor.push([result[1]]);
    passive.push([result[2]]);
    promoter.push([result[3]]); 
    total.push([result[4]]);     
  });

  summarySheet.getRange(startRow,3,offsetRow,1).setValues(total);
  summarySheet.getRange(startRow,4,offsetRow,1).setValues(detractor);
  summarySheet.getRange(startRow,5,offsetRow,1).setValues(passive);
  summarySheet.getRange(startRow,6,offsetRow,1).setValues(promoter);
  summarySheet.getRange(startRow,7,offsetRow,1).setValues(avgScore);
  summarySheet.getRange(startRow,7,offsetRow,1).setNumberFormat('0.00');
  summarySheet.getRange("B:G").setHorizontalAlignment('center');
  summarySheet.getRange("A1:G1").setBackground('#666666').setFontColor('#ffffff');
  summarySheet.autoResizeColumns(2, 6);
};
