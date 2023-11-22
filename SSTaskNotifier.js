function myFunction() 
{
    // Constants for task sequence (tsSeq)
    var tsSeqVertStartLine       = 2;   // Index of line where parsing starts
    var tsSeqVertLength          = 10;  // Count of lines that will be parsed
    var tsSeqTaskStatusLetter    = 'A'; // Letter of col that contains task status 
    var tsSeqTaskNameLetter      = 'B'; // Letter of col that contains task name
    var tsSeqTaskExecutorLetter  = 'C'; // Letter of col that contains name of executor
    var tsSeqDeadlineLetter      = 'D'; // Letter of col that contains the deadline date

    // Constants for Telegram sequence (tgSeq)
    var tgSeqVertStartLine      = 2;    // Index of line where parsing starts
    var tgSeqVertLength         = 5;    // Count of lines that will be parsed
    var tgSeqRecipientLetter    = 'G';  // Letter of col that contains recipient names
    var tgSeqUsernameLetter     = 'H';  // Letter of col that contain Telegram username (For using with @)

    // Service constants
    var tgBotKey                = "";   // Telegram Bot Token
    var tgMainChatId            = "";   // Telegram ChatId

    // Service variables
    var execs = "";
    var tgmap = {};
    var msgmap = {};
  
    //parsing TG sequence (tgSeq)
    for (var vpos = tgSeqVertStartLine; vpos < tgSeqVertStartLine+tgSeqVertLength; vpos++)
    {
        var user = SpreadsheetApp.getActiveSheet().getRange(tgSeqRecipientLetter+vpos).getValue();
        if (user != null && user != "")
        {
            var us = SpreadsheetApp.getActiveSheet().getRange(tgSeqRecipientLetter+vpos).getValue();
            var un = SpreadsheetApp.getActiveSheet().getRange(tgSeqUsernameLetter+vpos).getValue();
            tgmap[us] = un;
            msgmap[us] = "";
        }
    }
    
    // Parsing tasks sequence (tsSeq)
    for (var vpos1 = tsSeqVertStartLine; vpos1 < tsSeqVertStartLine+tsSeqVertLength; vpos1++)
    {
        var t = SpreadsheetApp.getActiveSheet().getRange(tsSeqTaskNameLetter+vpos1).getValue();
        var d = SpreadsheetApp.getActiveSheet().getRange(tsSeqDeadlineLetter+vpos1).getValue();
        var m = "(До " + d + ") " + t + "\n";
        if (SpreadsheetApp.getActiveSheet().getRange(tsSeqTaskStatusLetter+vpos1).getValue() == false && (t != null && t != ""))
        {
            var e = SpreadsheetApp.getActiveSheet().getRange(tsSeqTaskExecutorLetter+vpos1).getValue();
            if (e.includes(","))
            {
                var a = e.replace(" ", "").split(",");
                a.forEach((elem) => (msgmap[elem] != null && msgmap[elem] != undefined ? (msgmap[elem] += m) : (msgmap[elem] = m)));
                a.forEach((elem1) => (execs.includes(elem1) ? "" : execs += elem1 + ":"));
            }
            else
            {
                if (!execs.includes(e)) execs += e + ":";
                msgmap[e] != null && msgmap != undefined ? msgmap[e] += m : msgmap[e] = m;
            }
        }
    }

    // Building final message
    var f = "Активные задачи на текущий момент: \n\n";
    var s = execs.slice(0,-1).split(":");
    for (vpos2 = 0; vpos2 < s.length; vpos2++)
    {
        var x = tgmap[s[vpos2]] != null && tgmap[s[vpos2]] != undefined ? (s[vpos2] + " (@" + tgmap[s[vpos2]]) + "):" : s[vpos2];
        f += x + "\n" + msgmap[s[vpos2]] + "\n";
    }
    f += "\n LINT-TO-TAB"
    
    // Sending whole message to Telegram
    UrlFetchApp.fetch("https://api.telegram.org/bot"+tgBotKey+"/sendMessage", {
           'method' : 'post',
           'payload' : {
               chat_id: tgMainChatId,
               text: f
           }
       });
}