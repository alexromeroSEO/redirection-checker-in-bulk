function RedirectChecker() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ufila = ss.getLastRow();
  var DValores = ss.getRange("D3:D").getValues(); 
  var Dultima = DValores.filter(String).length;
  i = Dultima + 3
  var urls = ss
                    .getRange("B"+i+":B" + ufila)
                    .getValues()
                    .flat();
  options = { 
    'muteHttpExceptions': true,
    'followRedirects': false
  }
  
  const iterator = urls.forEach(url => {
    var IdealUrl = ss.getRange("C"+i+":C"+i).getValues() 
    var response = UrlFetchApp.fetch(url, options );
    var responseCode = response.getResponseCode()
    var redirect = /^3.+/.test(responseCode);
    if(redirect){
      ss.getRange("D"+i+":D"+i).setValue('HTTP '+responseCode)
      var FinalURL = response.getHeaders().Location;
      ss.getRange("E"+i+":E"+i).setValue(FinalURL)
        if(FinalURL == IdealUrl){
        ss.getRange("F"+i+":F"+i).setValue("Si")  
        }else{
        ss.getRange("F"+i+":F"+i).setValue("No") 
        var response2 = UrlFetchApp.fetch(FinalURL, options);
        var FinalURL2 = response2.getHeaders().Location;
        if(FinalURL2){
          ss.getRange("G"+i+":G"+i).setValue(FinalURL2)
          ss.getRange("K"+i+":K"+i).setValue('WARNING: Bucle de redirecciones').setBackground("#ffe599")
        if(FinalURL2 == IdealUrl){
        ss.getRange("H"+i+":H"+i).setValue("Si")  
        }else{
        ss.getRange("H"+i+":H"+i).setValue("No") 
        var response3 = UrlFetchApp.fetch(FinalURL2, options);
        var FinalURL3 = response3.getHeaders().Location;
        if(FinalURL3){
          ss.getRange("I"+i+":I"+i).setValue(FinalURL3)
          ss.getRange("J"+i+":J"+i).setValue(FinalURL3 == IdealUrl ? "Si" : "No")
        }else{}
      }
        }else{}
      }
    }else{
      ss.getRange("D"+i+":D"+i).setValue('HTTP '+responseCode)
    }
    i++
    }
    )
}

  
