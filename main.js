function fetchPage() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ufila = ss.getLastRow();
    i = 3;
    var urls = ss
                      .getRange("B3:B" + ufila)
                      .getValues()
                      .flat();
    options = {
      'method': 'post', 
      'muteHttpExceptions': true, 
      'followRedirects': false
    }
    
    const iterator = urls.forEach(url => {
      var IdealUrl = ss.getRange("C"+i+":C"+i).getValues() 
      var response = UrlFetchApp.fetch(url, options );
      var FinalURL = response.getHeaders().Location;
      var responseCode = response.getResponseCode()
      ss.getRange("F"+i+":F"+i).setValue(FinalURL == IdealUrl ? "Si" : "No")
      var redirect = /^3.+/.test(responseCode);
      ss.getRange("D"+i+":D"+i).setValue('HTTP '+responseCode)
      ss.getRange("E"+i+":E"+i).setValue(redirect  ? FinalURL : "")
      if(redirect){
        var response2 = UrlFetchApp.fetch(FinalURL, options );
        var FinalURL2 = response2.getHeaders().Location;
        var responseCode2 = response2.getResponseCode()
        ss.getRange("G"+i+":G"+i).setValue(FinalURL2)
        ss.getRange("H"+i+":H"+i).setValue(FinalURL2 == IdealUrl ? "Si" : "No")
        debugger
      }
      i++
      }
      )
  }
    
  
  
  