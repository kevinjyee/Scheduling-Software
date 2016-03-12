
  //Work Allocation updates -- this is needed because an automatic update (With link cells) will create errors.
   /* function generateWorkAlloc_(){

      
var ss = SpreadsheetApp.getActiveSpreadsheet();
      //Dates
  var source = ss.getRange('Data!A1');
    source.copyTo(ss.getRange('Week1!A2:F2'), {contentsOnly: true});
    var source = ss.getRange('Data!B1');
    source.copyTo(ss.getRange('Week1!G2:L2'), {contentsOnly: true});
    var source = ss.getRange('Data!C1');
    source.copyTo(ss.getRange('Week1!M2:R2'), {contentsOnly: true});
    var source = ss.getRange('Data!D1');
    source.copyTo(ss.getRange('Week1!S2:X2'), {contentsOnly: true});
    var source = ss.getRange('Data!E1');
    source.copyTo(ss.getRange('Week1!Y2:AD2'), {contentsOnly: true});
          var source = ss.getRange('Data!F1');
    source.copyTo(ss.getRange('Week1!AE2:AJ2'), {contentsOnly: true});
          var source = ss.getRange('Data!G1');
    source.copyTo(ss.getRange('Week1!AK2:AP2'), {contentsOnly: true});
        var source = ss.getRange('Data!H1');
    source.copyTo(ss.getRange('Week1!AQ2:AV2'), {contentsOnly: true});
      
  
      
      //Column 1
      var source = ss.getRange('Data!A3:A14');
    source.copyTo(ss.getRange('Week1!C10:C21'), {contentsOnly: true});
       var source = ss.getRange('Data!A16:A69');
    source.copyTo(ss.getRange('Week1!C22:C75'), {contentsOnly: true});
       var source = ss.getRange('Data!A71:A78');
    source.copyTo(ss.getRange('Week1!C76:C83'), {contentsOnly: true});
       var source = ss.getRange('Data!A82:A87');
    source.copyTo(ss.getRange('Week1!C84:C89'), {contentsOnly: true});
       var source = ss.getRange('Data!A89:A90');
    source.copyTo(ss.getRange('Week1!C90:C91'), {contentsOnly: true});
       var source = ss.getRange('Data!A92:A93');
    source.copyTo(ss.getRange('Week1!C92:C93'), {contentsOnly: true});
       var source = ss.getRange('Data!A95:A96');
    source.copyTo(ss.getRange('Week1!C94:C95'), {contentsOnly: true});
       var source = ss.getRange('Data!A98:A99');
    source.copyTo(ss.getRange('Week1!C96:C97'), {contentsOnly: true});
       var source = ss.getRange('Data!A101:A102');
    source.copyTo(ss.getRange('Week1!C98:C99'), {contentsOnly: true});

     
      //Column 2
      
      
         var source = ss.getRange('Data!B3:B14');
    source.copyTo(ss.getRange('Week1!I10:I21'), {contentsOnly: true});
       var source = ss.getRange('Data!B16:B69');
    source.copyTo(ss.getRange('Week1!I22:I75'), {contentsOnly: true});
       var source = ss.getRange('Data!B71:B78');
    source.copyTo(ss.getRange('Week1!I76:I83'), {contentsOnly: true});
       var source = ss.getRange('Data!B82:B87');
    source.copyTo(ss.getRange('Week1!I84:I89'), {contentsOnly: true});
       var source = ss.getRange('Data!B89:B90');
    source.copyTo(ss.getRange('Week1!I90:I91'), {contentsOnly: true});
       var source = ss.getRange('Data!B92:B93');
    source.copyTo(ss.getRange('Week1!I92:I93'), {contentsOnly: true});
       var source = ss.getRange('Data!B95:B96');
    source.copyTo(ss.getRange('Week1!I94:I95'), {contentsOnly: true});
       var source = ss.getRange('Data!B98:B99');
    source.copyTo(ss.getRange('Week1!I96:I97'), {contentsOnly: true});
       var source = ss.getRange('Data!B101:B102');
    source.copyTo(ss.getRange('Week1!I98:I99'), {contentsOnly: true});

      
      //Column 3
      
       var source = ss.getRange('Data!C3:C14');
    source.copyTo(ss.getRange('Week1!O10:O21'), {contentsOnly: true});
       var source = ss.getRange('Data!C16:C69');
    source.copyTo(ss.getRange('Week1!O22:O75'), {contentsOnly: true});
       var source = ss.getRange('Data!C71:C78');
    source.copyTo(ss.getRange('Week1!O76:O83'), {contentsOnly: true});
       var source = ss.getRange('Data!C82:C87');
    source.copyTo(ss.getRange('Week1!O84:O89'), {contentsOnly: true});
       var source = ss.getRange('Data!C89:C90');
    source.copyTo(ss.getRange('Week1!O90:O91'), {contentsOnly: true});
       var source = ss.getRange('Data!C92:C93');
    source.copyTo(ss.getRange('Week1!O92:O93'), {contentsOnly: true});
       var source = ss.getRange('Data!C95:C96');
    source.copyTo(ss.getRange('Week1!O94:O95'), {contentsOnly: true});
       var source = ss.getRange('Data!C98:C99');
    source.copyTo(ss.getRange('Week1!O96:O97'), {contentsOnly: true});
       var source = ss.getRange('Data!C101:C102');
    source.copyTo(ss.getRange('Week1!O98:O99'), {contentsOnly: true});

      
      //Column 4
         var source = ss.getRange('Data!D3:D14');
    source.copyTo(ss.getRange('Week1!U10:U21'), {contentsOnly: true});
       var source = ss.getRange('Data!D16:D69');
    source.copyTo(ss.getRange('Week1!U22:U75'), {contentsOnly: true});
       var source = ss.getRange('Data!D71:D78');
    source.copyTo(ss.getRange('Week1!U76:U83'), {contentsOnly: true});
       var source = ss.getRange('Data!D82:D87');
    source.copyTo(ss.getRange('Week1!U84:U89'), {contentsOnly: true});
       var source = ss.getRange('Data!D89:D90');
    source.copyTo(ss.getRange('Week1!U90:U91'), {contentsOnly: true});
       var source = ss.getRange('Data!D92:D93');
    source.copyTo(ss.getRange('Week1!U92:U93'), {contentsOnly: true});
       var source = ss.getRange('Data!D95:D96');
    source.copyTo(ss.getRange('Week1!U94:U95'), {contentsOnly: true});
       var source = ss.getRange('Data!D98:D99');
    source.copyTo(ss.getRange('Week1!U96:U97'), {contentsOnly: true});
       var source = ss.getRange('Data!D101:D102');
    source.copyTo(ss.getRange('Week1!U98:U99'), {contentsOnly: true});

      
      //Column 5
         var source = ss.getRange('Data!E3:E14');
    source.copyTo(ss.getRange('Week1!AA10:AA21'), {contentsOnly: true});
       var source = ss.getRange('Data!E16:E69');
    source.copyTo(ss.getRange('Week1!AA22:AA75'), {contentsOnly: true});
       var source = ss.getRange('Data!E71:E78');
    source.copyTo(ss.getRange('Week1!AA76:AA83'), {contentsOnly: true});
       var source = ss.getRange('Data!E82:E87');
    source.copyTo(ss.getRange('Week1!AA84:AA89'), {contentsOnly: true});
       var source = ss.getRange('Data!E89:E90');
    source.copyTo(ss.getRange('Week1!AA90:AA91'), {contentsOnly: true});
       var source = ss.getRange('Data!E92:E93');
    source.copyTo(ss.getRange('Week1!AA92:AA93'), {contentsOnly: true});
       var source = ss.getRange('Data!E95:E96');
    source.copyTo(ss.getRange('Week1!AA94:AA95'), {contentsOnly: true});
       var source = ss.getRange('Data!E98:E99');
    source.copyTo(ss.getRange('Week1!AA96:AA97'), {contentsOnly: true});
       var source = ss.getRange('Data!E101:E102');
    source.copyTo(ss.getRange('Week1!AA98:AA99'), {contentsOnly: true});

      //Column 6
         var source = ss.getRange('Data!F3:F14');
    source.copyTo(ss.getRange('Week1!AG10:AG21'), {contentsOnly: true});
       var source = ss.getRange('Data!F16:F69');
    source.copyTo(ss.getRange('Week1!AG22:AG75'), {contentsOnly: true});
       var source = ss.getRange('Data!F71:F78');
    source.copyTo(ss.getRange('Week1!AG76:AG83'), {contentsOnly: true});
       var source = ss.getRange('Data!F82:F87');
    source.copyTo(ss.getRange('Week1!AG84:AG89'), {contentsOnly: true});
       var source = ss.getRange('Data!F89:F90');
    source.copyTo(ss.getRange('Week1!AG90:AG91'), {contentsOnly: true});
       var source = ss.getRange('Data!F92:F93');
    source.copyTo(ss.getRange('Week1!AG92:AG93'), {contentsOnly: true});
       var source = ss.getRange('Data!F95:F96');
    source.copyTo(ss.getRange('Week1!AG94:AG95'), {contentsOnly: true});
       var source = ss.getRange('Data!F98:F99');
    source.copyTo(ss.getRange('Week1!AG96:AG97'), {contentsOnly: true});
       var source = ss.getRange('Data!F101:F102');
    source.copyTo(ss.getRange('Week1!AG98:AG99'), {contentsOnly: true});

      
      //Column 7
      
         var source = ss.getRange('Data!G3:G14');
    source.copyTo(ss.getRange('Week1!AM10:AM21'), {contentsOnly: true});
       var source = ss.getRange('Data!G16:G69');
    source.copyTo(ss.getRange('Week1!AM22:AM75'), {contentsOnly: true});
       var source = ss.getRange('Data!G71:G78');
    source.copyTo(ss.getRange('Week1!AM76:AM83'), {contentsOnly: true});
       var source = ss.getRange('Data!G82:G87');
    source.copyTo(ss.getRange('Week1!AM84:AM89'), {contentsOnly: true});
       var source = ss.getRange('Data!G89:G90');
    source.copyTo(ss.getRange('Week1!AM90:AM91'), {contentsOnly: true});
       var source = ss.getRange('Data!G92:G93');
    source.copyTo(ss.getRange('Week1!AM92:AM93'), {contentsOnly: true});
       var source = ss.getRange('Data!G95:G96');
    source.copyTo(ss.getRange('Week1!AM94:AM95'), {contentsOnly: true});
       var source = ss.getRange('Data!G98:G99');
    source.copyTo(ss.getRange('Week1!AM96:AM97'), {contentsOnly: true});
       var source = ss.getRange('Data!G101:G102');
    source.copyTo(ss.getRange('Week1!AM98:AM99'), {contentsOnly: true});

      //Column 8
      
         var source = ss.getRange('Data!H3:H14');
    source.copyTo(ss.getRange('Week1!AS10:AS21'), {contentsOnly: true});
       var source = ss.getRange('Data!H16:H69');
    source.copyTo(ss.getRange('Week1!AS22:AS75'), {contentsOnly: true});
       var source = ss.getRange('Data!H71:H78');
    source.copyTo(ss.getRange('Week1!AS76:AS83'), {contentsOnly: true});
       var source = ss.getRange('Data!H82:H87');
    source.copyTo(ss.getRange('Week1!AS84:AS89'), {contentsOnly: true});
       var source = ss.getRange('Data!H89:H90');
    source.copyTo(ss.getRange('Week1!AS90:AS91'), {contentsOnly: true});
       var source = ss.getRange('Data!H92:H93');
    source.copyTo(ss.getRange('Week1!AS92:AS93'), {contentsOnly: true});
       var source = ss.getRange('Data!H95:H96');
    source.copyTo(ss.getRange('Week1!AS94:AS95'), {contentsOnly: true});
       var source = ss.getRange('Data!H98:H99');
    source.copyTo(ss.getRange('Week1!AS96:AS97'), {contentsOnly: true});
       var source = ss.getRange('Data!H101:H102');
    source.copyTo(ss.getRange('Week1!AS98:AS99'), {contentsOnly: true});

    



/////////////////Molecules/////////////////////////////////
//Column 1

   var source = ss.getRange('Molecule!A3:A14');
    source.copyTo(ss.getRange('Week1!B10:B21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!A16:A69');
    source.copyTo(ss.getRange('Week1!B22:B75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!A71:A78');
    source.copyTo(ss.getRange('Week1!B76:B83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!A82:A87');
    source.copyTo(ss.getRange('Week1!B84:B89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!A89:A90');
    source.copyTo(ss.getRange('Week1!B90:B91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!A92:A93');
    source.copyTo(ss.getRange('Week1!B92:B93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!A95:A96');
    source.copyTo(ss.getRange('Week1!B94:B95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!A98:A99');
    source.copyTo(ss.getRange('Week1!B96:B97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!A101:A102');
    source.copyTo(ss.getRange('Week1!B98:B99'), {contentsOnly: true});
//Column 2
   var source = ss.getRange('Molecule!B3:B14');
    source.copyTo(ss.getRange('Week1!H10:H21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!B16:B69');
    source.copyTo(ss.getRange('Week1!H22:H75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!B71:B78');
    source.copyTo(ss.getRange('Week1!H76:H83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!B82:B87');
    source.copyTo(ss.getRange('Week1!H84:H89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!B89:B90');
    source.copyTo(ss.getRange('Week1!H90:H91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!B92:B93');
    source.copyTo(ss.getRange('Week1!H92:H93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!B95:B96');
    source.copyTo(ss.getRange('Week1!H94:H95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!B98:B99');
    source.copyTo(ss.getRange('Week1!H96:H97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!B101:B102');
    source.copyTo(ss.getRange('Week1!H98:H99'), {contentsOnly: true});

//Column 3
   var source = ss.getRange('Molecule!C3:C14');
    source.copyTo(ss.getRange('Week1!N10:N21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!C16:C69');
    source.copyTo(ss.getRange('Week1!N22:N75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!C71:C78');
    source.copyTo(ss.getRange('Week1!N76:N83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!C82:C87');
    source.copyTo(ss.getRange('Week1!N84:N89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!C89:C90');
    source.copyTo(ss.getRange('Week1!N90:N91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!C92:C93');
    source.copyTo(ss.getRange('Week1!N92:N93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!C95:C96');
    source.copyTo(ss.getRange('Week1!N94:N95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!C98:C99');
    source.copyTo(ss.getRange('Week1!N96:N97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!C101:C102');
    source.copyTo(ss.getRange('Week1!N98:N99'), {contentsOnly: true});

//Column 4
   var source = ss.getRange('Molecule!D3:D14');
    source.copyTo(ss.getRange('Week1!T10:T21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!D16:D69');
    source.copyTo(ss.getRange('Week1!T22:T75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!D71:D78');
    source.copyTo(ss.getRange('Week1!T76:T83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!D82:D87');
    source.copyTo(ss.getRange('Week1!T84:T89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!D89:D90');
    source.copyTo(ss.getRange('Week1!T90:T91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!D92:D93');
    source.copyTo(ss.getRange('Week1!T92:T93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!D95:D96');
    source.copyTo(ss.getRange('Week1!T94:T95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!D98:D99');
    source.copyTo(ss.getRange('Week1!T96:T97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!D101:D102');
    source.copyTo(ss.getRange('Week1!T98:T99'), {contentsOnly: true});

//Column 5
   var source = ss.getRange('Molecule!E3:E14');
    source.copyTo(ss.getRange('Week1!Z10:Z21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!E16:E69');
    source.copyTo(ss.getRange('Week1!Z22:Z75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!E71:E78');
    source.copyTo(ss.getRange('Week1!Z76:Z83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!E82:E87');
    source.copyTo(ss.getRange('Week1!Z84:Z89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!E89:E90');
    source.copyTo(ss.getRange('Week1!Z90:Z91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!E92:E93');
    source.copyTo(ss.getRange('Week1!Z92:Z93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!E95:E96');
    source.copyTo(ss.getRange('Week1!Z94:Z95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!E98:E99');
    source.copyTo(ss.getRange('Week1!Z96:Z97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!E101:E102');
    source.copyTo(ss.getRange('Week1!Z98:Z99'), {contentsOnly: true});

//Column 6
   var source = ss.getRange('Molecule!F3:F14');
    source.copyTo(ss.getRange('Week1!AF10:AF21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!F16:F69');
    source.copyTo(ss.getRange('Week1!AF22:AF75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!F71:F78');
    source.copyTo(ss.getRange('Week1!AF76:AF83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!F82:F87');
    source.copyTo(ss.getRange('Week1!AF84:AF89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!F89:F90');
    source.copyTo(ss.getRange('Week1!AF90:AF91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!F92:F93');
    source.copyTo(ss.getRange('Week1!AF92:AF93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!F95:F96');
    source.copyTo(ss.getRange('Week1!AF94:AF95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!F98:F99');
    source.copyTo(ss.getRange('Week1!AF96:AF97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!F101:F102');
    source.copyTo(ss.getRange('Week1!AF98:AF99'), {contentsOnly: true});

//Column 7
   var source = ss.getRange('Molecule!G3:G14');
    source.copyTo(ss.getRange('Week1!AL10:AL21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!G16:G69');
    source.copyTo(ss.getRange('Week1!AL22:AL75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!G71:G78');
    source.copyTo(ss.getRange('Week1!AL76:AL83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!G82:G87');
    source.copyTo(ss.getRange('Week1!AL84:AL89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!G89:G90');
    source.copyTo(ss.getRange('Week1!AL90:AL91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!G92:G93');
    source.copyTo(ss.getRange('Week1!AL92:AL93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!G95:G96');
    source.copyTo(ss.getRange('Week1!AL94:AL95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!G98:G99');
    source.copyTo(ss.getRange('Week1!AL96:AL97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!G101:G102');
    source.copyTo(ss.getRange('Week1!AL98:AL99'), {contentsOnly: true});

//Column 8
   var source = ss.getRange('Molecule!H3:H14');
    source.copyTo(ss.getRange('Week1!AR10:AR21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!H16:H69');
    source.copyTo(ss.getRange('Week1!AR22:AR75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!H71:H78');
    source.copyTo(ss.getRange('Week1!AR76:AR83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!H82:H87');
    source.copyTo(ss.getRange('Week1!AR84:AR89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!H89:H90');
    source.copyTo(ss.getRange('Week1!AR90:AR91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!H92:H93');
    source.copyTo(ss.getRange('Week1!AR92:AR93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!H95:H96');
    source.copyTo(ss.getRange('Week1!AR94:AR95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!H98:H99');
    source.copyTo(ss.getRange('Week1!AR96:AR97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!H101:H102');
    source.copyTo(ss.getRange('Week1!AR98:AR99'), {contentsOnly: true});

      ///---_Week 2//////////////
      
      
      //Dates
  var source = ss.getRange('Data!H1');
    source.copyTo(ss.getRange('Week2!A2:F2'), {contentsOnly: true});
    var source = ss.getRange('Data!I1');
    source.copyTo(ss.getRange('Week2!G2:L2'), {contentsOnly: true});
    var source = ss.getRange('Data!J1');
    source.copyTo(ss.getRange('Week2!M2:R2'), {contentsOnly: true});
    var source = ss.getRange('Data!K1');
    source.copyTo(ss.getRange('Week2!S2:X2'), {contentsOnly: true});
    var source = ss.getRange('Data!L1');
    source.copyTo(ss.getRange('Week2!Y2:AD2'), {contentsOnly: true});
        var source = ss.getRange('Data!M1');
    source.copyTo(ss.getRange('Week2!AE2:AJ2'), {contentsOnly: true});
          var source = ss.getRange('Data!N1');
    source.copyTo(ss.getRange('Week2!AK2:AP2'), {contentsOnly: true});
        var source = ss.getRange('Data!O1');
    source.copyTo(ss.getRange('Week2!AQ2:AV2'), {contentsOnly: true});

//Column 1
         var source = ss.getRange('Data!H3:H14');
    source.copyTo(ss.getRange('Week2!C10:C21'), {contentsOnly: true});
       var source = ss.getRange('Data!H16:H69');
    source.copyTo(ss.getRange('Week2!C22:C75'), {contentsOnly: true});
       var source = ss.getRange('Data!H71:H78');
    source.copyTo(ss.getRange('Week2!C76:C83'), {contentsOnly: true});
       var source = ss.getRange('Data!H82:H87');
    source.copyTo(ss.getRange('Week2!C84:C89'), {contentsOnly: true});
       var source = ss.getRange('Data!H89:H90');
    source.copyTo(ss.getRange('Week2!C90:C91'), {contentsOnly: true});
       var source = ss.getRange('Data!H92:H93');
    source.copyTo(ss.getRange('Week2!C92:C93'), {contentsOnly: true});
       var source = ss.getRange('Data!H95:H96');
    source.copyTo(ss.getRange('Week2!C94:C95'), {contentsOnly: true});
       var source = ss.getRange('Data!H98:H99');
    source.copyTo(ss.getRange('Week2!C96:C97'), {contentsOnly: true});
       var source = ss.getRange('Data!H101:H102');
    source.copyTo(ss.getRange('Week2!C98:C99'), {contentsOnly: true});

 //Column 2
         var source = ss.getRange('Data!I3:I14');
    source.copyTo(ss.getRange('Week2!I10:I21'), {contentsOnly: true});
       var source = ss.getRange('Data!I16:I69');
    source.copyTo(ss.getRange('Week2!I22:I75'), {contentsOnly: true});
       var source = ss.getRange('Data!I71:I78');
    source.copyTo(ss.getRange('Week2!I76:I83'), {contentsOnly: true});
       var source = ss.getRange('Data!I82:I87');
    source.copyTo(ss.getRange('Week2!I84:I89'), {contentsOnly: true});
       var source = ss.getRange('Data!I89:I90');
    source.copyTo(ss.getRange('Week2!I90:I91'), {contentsOnly: true});
       var source = ss.getRange('Data!I92:I93');
    source.copyTo(ss.getRange('Week2!I92:I93'), {contentsOnly: true});
       var source = ss.getRange('Data!I95:I96');
    source.copyTo(ss.getRange('Week2!I94:I95'), {contentsOnly: true});
       var source = ss.getRange('Data!I98:I99');
    source.copyTo(ss.getRange('Week2!I96:I97'), {contentsOnly: true});
       var source = ss.getRange('Data!I101:I102');
    source.copyTo(ss.getRange('Week2!I98:I99'), {contentsOnly: true});

      //Column 3
         var source = ss.getRange('Data!J3:J14');
    source.copyTo(ss.getRange('Week2!O10:O21'), {contentsOnly: true});
       var source = ss.getRange('Data!J16:J69');
    source.copyTo(ss.getRange('Week2!O22:O75'), {contentsOnly: true});
       var source = ss.getRange('Data!J71:J78');
    source.copyTo(ss.getRange('Week2!O76:O83'), {contentsOnly: true});
       var source = ss.getRange('Data!J82:J87');
    source.copyTo(ss.getRange('Week2!O84:O89'), {contentsOnly: true});
       var source = ss.getRange('Data!J89:J90');
    source.copyTo(ss.getRange('Week2!O90:O91'), {contentsOnly: true});
       var source = ss.getRange('Data!J92:J93');
    source.copyTo(ss.getRange('Week2!O92:O93'), {contentsOnly: true});
       var source = ss.getRange('Data!J95:J96');
    source.copyTo(ss.getRange('Week2!O94:O95'), {contentsOnly: true});
       var source = ss.getRange('Data!J98:J99');
    source.copyTo(ss.getRange('Week2!O96:O97'), {contentsOnly: true});
       var source = ss.getRange('Data!J101:J102');
    source.copyTo(ss.getRange('Week2!O98:O99'), {contentsOnly: true});

      //Column 4
         var source = ss.getRange('Data!K3:K14');
    source.copyTo(ss.getRange('Week2!U10:U21'), {contentsOnly: true});
       var source = ss.getRange('Data!K16:K69');
    source.copyTo(ss.getRange('Week2!U22:U75'), {contentsOnly: true});
       var source = ss.getRange('Data!K71:K78');
    source.copyTo(ss.getRange('Week2!U76:U83'), {contentsOnly: true});
       var source = ss.getRange('Data!K82:K87');
    source.copyTo(ss.getRange('Week2!U84:U89'), {contentsOnly: true});
       var source = ss.getRange('Data!K89:K90');
    source.copyTo(ss.getRange('Week2!U90:U91'), {contentsOnly: true});
       var source = ss.getRange('Data!K92:K93');
    source.copyTo(ss.getRange('Week2!U92:U93'), {contentsOnly: true});
       var source = ss.getRange('Data!K95:K96');
    source.copyTo(ss.getRange('Week2!U94:U95'), {contentsOnly: true});
       var source = ss.getRange('Data!K98:K99');
    source.copyTo(ss.getRange('Week2!U96:U97'), {contentsOnly: true});
       var source = ss.getRange('Data!K101:K102');
    source.copyTo(ss.getRange('Week2!U98:U99'), {contentsOnly: true});

      //Column 5
         var source = ss.getRange('Data!L3:L14');
    source.copyTo(ss.getRange('Week2!AA10:AA21'), {contentsOnly: true});
       var source = ss.getRange('Data!L16:L69');
    source.copyTo(ss.getRange('Week2!AA22:AA75'), {contentsOnly: true});
       var source = ss.getRange('Data!L71:L78');
    source.copyTo(ss.getRange('Week2!AA76:AA83'), {contentsOnly: true});
       var source = ss.getRange('Data!L82:L87');
    source.copyTo(ss.getRange('Week2!AA84:AA89'), {contentsOnly: true});
       var source = ss.getRange('Data!L89:L90');
    source.copyTo(ss.getRange('Week2!AA90:AA91'), {contentsOnly: true});
       var source = ss.getRange('Data!L92:L93');
    source.copyTo(ss.getRange('Week2!AA92:AA93'), {contentsOnly: true});
       var source = ss.getRange('Data!L95:L96');
    source.copyTo(ss.getRange('Week2!AA94:AA95'), {contentsOnly: true});
       var source = ss.getRange('Data!L98:L99');
    source.copyTo(ss.getRange('Week2!AA96:AA97'), {contentsOnly: true});
       var source = ss.getRange('Data!L101:L102');
    source.copyTo(ss.getRange('Week2!AA98:AA99'), {contentsOnly: true});

      //Column 6
         var source = ss.getRange('Data!M3:M14');
    source.copyTo(ss.getRange('Week2!AG10:AG21'), {contentsOnly: true});
       var source = ss.getRange('Data!M16:M69');
    source.copyTo(ss.getRange('Week2!AG22:AG75'), {contentsOnly: true});
       var source = ss.getRange('Data!M71:M78');
    source.copyTo(ss.getRange('Week2!AG76:AG83'), {contentsOnly: true});
       var source = ss.getRange('Data!M82:M87');
    source.copyTo(ss.getRange('Week2!AG84:AG89'), {contentsOnly: true});
       var source = ss.getRange('Data!M89:M90');
    source.copyTo(ss.getRange('Week2!AG90:AG91'), {contentsOnly: true});
       var source = ss.getRange('Data!M92:M93');
    source.copyTo(ss.getRange('Week2!AG92:AG93'), {contentsOnly: true});
       var source = ss.getRange('Data!M95:M96');
    source.copyTo(ss.getRange('Week2!AG94:AG95'), {contentsOnly: true});
       var source = ss.getRange('Data!M98:M99');
    source.copyTo(ss.getRange('Week2!AG96:AG97'), {contentsOnly: true});
       var source = ss.getRange('Data!M101:M102');
    source.copyTo(ss.getRange('Week2!AG98:AG99'), {contentsOnly: true});

      //Column 7
         var source = ss.getRange('Data!N3:N14');
    source.copyTo(ss.getRange('Week2!AM10:AM21'), {contentsOnly: true});
       var source = ss.getRange('Data!N16:N69');
    source.copyTo(ss.getRange('Week2!AM22:AM75'), {contentsOnly: true});
       var source = ss.getRange('Data!N71:N78');
    source.copyTo(ss.getRange('Week2!AM76:AM83'), {contentsOnly: true});
       var source = ss.getRange('Data!N82:N87');
    source.copyTo(ss.getRange('Week2!AM84:AM89'), {contentsOnly: true});
       var source = ss.getRange('Data!N89:N90');
    source.copyTo(ss.getRange('Week2!AM90:AM91'), {contentsOnly: true});
       var source = ss.getRange('Data!N92:N93');
    source.copyTo(ss.getRange('Week2!AM92:AM93'), {contentsOnly: true});
       var source = ss.getRange('Data!N95:N96');
    source.copyTo(ss.getRange('Week2!AM94:AM95'), {contentsOnly: true});
       var source = ss.getRange('Data!N98:N99');
    source.copyTo(ss.getRange('Week2!AM96:AM97'), {contentsOnly: true});
       var source = ss.getRange('Data!N101:N102');
    source.copyTo(ss.getRange('Week2!AM98:AM99'), {contentsOnly: true});

      //Column 8
         var source = ss.getRange('Data!O3:O14');
    source.copyTo(ss.getRange('Week2!AS10:AS21'), {contentsOnly: true});
       var source = ss.getRange('Data!O16:O69');
    source.copyTo(ss.getRange('Week2!AS22:AS75'), {contentsOnly: true});
       var source = ss.getRange('Data!O71:O78');
    source.copyTo(ss.getRange('Week2!AS76:AS83'), {contentsOnly: true});
       var source = ss.getRange('Data!O82:O87');
    source.copyTo(ss.getRange('Week2!AS84:AS89'), {contentsOnly: true});
       var source = ss.getRange('Data!O89:O90');
    source.copyTo(ss.getRange('Week2!AS90:AS91'), {contentsOnly: true});
       var source = ss.getRange('Data!O92:O93');
    source.copyTo(ss.getRange('Week2!AS92:AS93'), {contentsOnly: true});
       var source = ss.getRange('Data!O95:O96');
    source.copyTo(ss.getRange('Week2!AS94:AS95'), {contentsOnly: true});
       var source = ss.getRange('Data!O98:O99');
    source.copyTo(ss.getRange('Week2!AS96:AS97'), {contentsOnly: true});
       var source = ss.getRange('Data!O101:O102');
    source.copyTo(ss.getRange('Week2!AS98:AS99'), {contentsOnly: true});


      /////Molecule////////////////////
      //Column 1
         var source = ss.getRange('Molecule!H3:H14');
    source.copyTo(ss.getRange('Week2!B10:B21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!H16:H69');
    source.copyTo(ss.getRange('Week2!B22:B75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!H71:H78');
    source.copyTo(ss.getRange('Week2!B76:B83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!H82:H87');
    source.copyTo(ss.getRange('Week2!B84:B89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!H89:H90');
    source.copyTo(ss.getRange('Week2!B90:B91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!H92:H93');
    source.copyTo(ss.getRange('Week2!B92:B93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!H95:H96');
    source.copyTo(ss.getRange('Week2!B94:B95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!H98:H99');
    source.copyTo(ss.getRange('Week2!B96:B97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!H101:H102');
    source.copyTo(ss.getRange('Week2!B98:B99'), {contentsOnly: true});

      //Column 2
         var source = ss.getRange('Molecule!I3:I14');
    source.copyTo(ss.getRange('Week2!H10:H21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!I16:I69');
    source.copyTo(ss.getRange('Week2!H22:H75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!I71:I78');
    source.copyTo(ss.getRange('Week2!H76:H83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!I82:I87');
    source.copyTo(ss.getRange('Week2!H84:H89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!I89:I90');
    source.copyTo(ss.getRange('Week2!H90:H91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!I92:I93');
    source.copyTo(ss.getRange('Week2!H92:H93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!I95:I96');
    source.copyTo(ss.getRange('Week2!H94:H95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!I98:I99');
    source.copyTo(ss.getRange('Week2!H96:H97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!I101:I102');
    source.copyTo(ss.getRange('Week2!H98:H99'), {contentsOnly: true});

      //Column 3
         var source = ss.getRange('Molecule!J3:J14');
    source.copyTo(ss.getRange('Week2!N10:N21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!J16:J69');
    source.copyTo(ss.getRange('Week2!N22:N75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!J71:J78');
    source.copyTo(ss.getRange('Week2!N76:N83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!J82:J87');
    source.copyTo(ss.getRange('Week2!N84:N89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!J89:J90');
    source.copyTo(ss.getRange('Week2!N90:N91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!J92:J93');
    source.copyTo(ss.getRange('Week2!N92:N93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!J95:J96');
    source.copyTo(ss.getRange('Week2!N94:N95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!J98:J99');
    source.copyTo(ss.getRange('Week2!N96:N97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!J101:J102');
    source.copyTo(ss.getRange('Week2!N98:N99'), {contentsOnly: true});

      //Column 4
         var source = ss.getRange('Molecule!K3:K14');
    source.copyTo(ss.getRange('Week2!T10:T21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!K16:K69');
    source.copyTo(ss.getRange('Week2!T22:T75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!K71:K78');
    source.copyTo(ss.getRange('Week2!T76:T83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!K82:K87');
    source.copyTo(ss.getRange('Week2!T84:T89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!K89:K90');
    source.copyTo(ss.getRange('Week2!T90:T91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!K92:K93');
    source.copyTo(ss.getRange('Week2!T92:T93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!K95:K96');
    source.copyTo(ss.getRange('Week2!T94:T95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!K98:K99');
    source.copyTo(ss.getRange('Week2!T96:T97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!K101:K102');
    source.copyTo(ss.getRange('Week2!T98:T99'), {contentsOnly: true});

      //Column 5
         var source = ss.getRange('Molecule!L3:L14');
    source.copyTo(ss.getRange('Week2!Z10:Z21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!L16:L69');
    source.copyTo(ss.getRange('Week2!Z22:Z75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!L71:L78');
    source.copyTo(ss.getRange('Week2!Z76:Z83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!L82:L87');
    source.copyTo(ss.getRange('Week2!Z84:Z89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!L89:L90');
    source.copyTo(ss.getRange('Week2!Z90:Z91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!L92:L93');
    source.copyTo(ss.getRange('Week2!Z92:Z93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!L95:L96');
    source.copyTo(ss.getRange('Week2!Z94:Z95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!L98:L99');
    source.copyTo(ss.getRange('Week2!Z96:Z97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!L101:L102');
    source.copyTo(ss.getRange('Week2!Z98:Z99'), {contentsOnly: true});

      //Column 6
         var source = ss.getRange('Molecule!M3:M14');
    source.copyTo(ss.getRange('Week2!AF10:AF21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!M16:M69');
    source.copyTo(ss.getRange('Week2!AF22:AF75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!M71:M78');
    source.copyTo(ss.getRange('Week2!AF76:AF83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!M82:M87');
    source.copyTo(ss.getRange('Week2!AF84:AF89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!M89:M90');
    source.copyTo(ss.getRange('Week2!AF90:AF91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!M92:M93');
    source.copyTo(ss.getRange('Week2!AF92:AF93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!M95:M96');
    source.copyTo(ss.getRange('Week2!AF94:AF95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!M98:M99');
    source.copyTo(ss.getRange('Week2!AF96:AF97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!M101:M102');
    source.copyTo(ss.getRange('Week2!AF98:AF99'), {contentsOnly: true});

      //Column 7
         var source = ss.getRange('Molecule!N3:N14');
    source.copyTo(ss.getRange('Week2!AL10:AL21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!N16:N69');
    source.copyTo(ss.getRange('Week2!AL22:AL75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!N71:N78');
    source.copyTo(ss.getRange('Week2!AL76:AL83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!N82:N87');
    source.copyTo(ss.getRange('Week2!AL84:AL89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!N89:N90');
    source.copyTo(ss.getRange('Week2!AL90:AL91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!N92:N93');
    source.copyTo(ss.getRange('Week2!AL92:AL93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!N95:N96');
    source.copyTo(ss.getRange('Week2!AL94:AL95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!N98:N99');
    source.copyTo(ss.getRange('Week2!AL96:AL97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!N101:N102');
    source.copyTo(ss.getRange('Week2!AL98:AL99'), {contentsOnly: true});

      
      //Column 8
         var source = ss.getRange('Molecule!O3:O14');
    source.copyTo(ss.getRange('Week2!AR10:AR21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!O16:O69');
    source.copyTo(ss.getRange('Week2!AR22:AR75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!O71:O78');
    source.copyTo(ss.getRange('Week2!AR76:AR83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!O82:O87');
    source.copyTo(ss.getRange('Week2!AR84:AR89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!O89:O90');
    source.copyTo(ss.getRange('Week2!AR90:AR91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!O92:O93');
    source.copyTo(ss.getRange('Week2!AR92:AR93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!O95:O96');
    source.copyTo(ss.getRange('Week2!AR94:AR95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!O98:O99');
    source.copyTo(ss.getRange('Week2!AR96:AR97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!O101:O102');
    source.copyTo(ss.getRange('Week2!AR98:AR99'), {contentsOnly: true});
      
      //Week 3
      //Dates
  var source = ss.getRange('Data!O1');
    source.copyTo(ss.getRange('Week3!A2:F2'), {contentsOnly: true});
    var source = ss.getRange('Data!P1');
    source.copyTo(ss.getRange('Week3!G2:L2'), {contentsOnly: true});
    var source = ss.getRange('Data!Q1');
    source.copyTo(ss.getRange('Week3!M2:R2'), {contentsOnly: true});
    var source = ss.getRange('Data!R1');
    source.copyTo(ss.getRange('Week3!S2:X2'), {contentsOnly: true});
    var source = ss.getRange('Data!S1');
    source.copyTo(ss.getRange('Week3!Y2:AD2'), {contentsOnly: true});
          var source = ss.getRange('Data!T1');
    source.copyTo(ss.getRange('Week3!AE2:AJ2'), {contentsOnly: true});
          var source = ss.getRange('Data!U1');
    source.copyTo(ss.getRange('Week3!AK2:AP2'), {contentsOnly: true});
        var source = ss.getRange('Data!V1');
    source.copyTo(ss.getRange('Week3!AQ2:AV2'), {contentsOnly: true});


      //Column 1
            var source = ss.getRange('Data!O3:O14');
    source.copyTo(ss.getRange('Week3!C10:C21'), {contentsOnly: true});
       var source = ss.getRange('Data!O16:O69');
    source.copyTo(ss.getRange('Week3!C22:C75'), {contentsOnly: true});
       var source = ss.getRange('Data!O71:O78');
    source.copyTo(ss.getRange('Week3!C76:C83'), {contentsOnly: true});
       var source = ss.getRange('Data!O82:O87');
    source.copyTo(ss.getRange('Week3!C84:C89'), {contentsOnly: true});
       var source = ss.getRange('Data!O89:O90');
    source.copyTo(ss.getRange('Week3!C90:C91'), {contentsOnly: true});
       var source = ss.getRange('Data!O92:O93');
    source.copyTo(ss.getRange('Week3!C92:C93'), {contentsOnly: true});
       var source = ss.getRange('Data!O95:O96');
    source.copyTo(ss.getRange('Week3!C94:C95'), {contentsOnly: true});
       var source = ss.getRange('Data!O98:O99');
    source.copyTo(ss.getRange('Week3!C96:C97'), {contentsOnly: true});
       var source = ss.getRange('Data!O101:O102');
    source.copyTo(ss.getRange('Week3!C98:C99'), {contentsOnly: true});

      //Column 2
            var source = ss.getRange('Data!P3:P14');
    source.copyTo(ss.getRange('Week3!I10:I21'), {contentsOnly: true});
       var source = ss.getRange('Data!P16:P69');
    source.copyTo(ss.getRange('Week3!I22:I75'), {contentsOnly: true});
       var source = ss.getRange('Data!P71:P78');
    source.copyTo(ss.getRange('Week3!I76:I83'), {contentsOnly: true});
       var source = ss.getRange('Data!P82:P87');
    source.copyTo(ss.getRange('Week3!I84:I89'), {contentsOnly: true});
       var source = ss.getRange('Data!P89:P90');
    source.copyTo(ss.getRange('Week3!I90:I91'), {contentsOnly: true});
       var source = ss.getRange('Data!P92:P93');
    source.copyTo(ss.getRange('Week3!I92:I93'), {contentsOnly: true});
       var source = ss.getRange('Data!P95:P96');
    source.copyTo(ss.getRange('Week3!I94:I95'), {contentsOnly: true});
       var source = ss.getRange('Data!P98:P99');
    source.copyTo(ss.getRange('Week3!I96:I97'), {contentsOnly: true});
       var source = ss.getRange('Data!P101:P102');
    source.copyTo(ss.getRange('Week3!I98:I99'), {contentsOnly: true});

      //Column 3
            var source = ss.getRange('Data!Q3:Q14');
    source.copyTo(ss.getRange('Week3!O10:O21'), {contentsOnly: true});
       var source = ss.getRange('Data!Q16:Q69');
    source.copyTo(ss.getRange('Week3!O22:O75'), {contentsOnly: true});
       var source = ss.getRange('Data!Q71:Q78');
    source.copyTo(ss.getRange('Week3!O76:O83'), {contentsOnly: true});
       var source = ss.getRange('Data!Q82:Q87');
    source.copyTo(ss.getRange('Week3!O84:O89'), {contentsOnly: true});
       var source = ss.getRange('Data!Q89:Q90');
    source.copyTo(ss.getRange('Week3!O90:O91'), {contentsOnly: true});
       var source = ss.getRange('Data!Q92:Q93');
    source.copyTo(ss.getRange('Week3!O92:O93'), {contentsOnly: true});
       var source = ss.getRange('Data!Q95:Q96');
    source.copyTo(ss.getRange('Week3!O94:O95'), {contentsOnly: true});
       var source = ss.getRange('Data!Q98:Q99');
    source.copyTo(ss.getRange('Week3!O96:O97'), {contentsOnly: true});
       var source = ss.getRange('Data!Q101:Q102');
    source.copyTo(ss.getRange('Week3!O98:O99'), {contentsOnly: true});

      //Column 4
          var source = ss.getRange('Data!R3:R14');
    source.copyTo(ss.getRange('Week3!U10:U21'), {contentsOnly: true});
       var source = ss.getRange('Data!R16:R69');
    source.copyTo(ss.getRange('Week3!U22:U75'), {contentsOnly: true});
       var source = ss.getRange('Data!R71:R78');
    source.copyTo(ss.getRange('Week3!U76:U83'), {contentsOnly: true});
       var source = ss.getRange('Data!R82:R87');
    source.copyTo(ss.getRange('Week3!U84:U89'), {contentsOnly: true});
       var source = ss.getRange('Data!R89:R90');
    source.copyTo(ss.getRange('Week3!U90:U91'), {contentsOnly: true});
       var source = ss.getRange('Data!R92:R93');
    source.copyTo(ss.getRange('Week3!U92:U93'), {contentsOnly: true});
       var source = ss.getRange('Data!R95:R96');
    source.copyTo(ss.getRange('Week3!U94:U95'), {contentsOnly: true});
       var source = ss.getRange('Data!R98:R99');
    source.copyTo(ss.getRange('Week3!U96:U97'), {contentsOnly: true});
       var source = ss.getRange('Data!R101:R102');
    source.copyTo(ss.getRange('Week3!U98:U99'), {contentsOnly: true});

      //Column 5
            var source = ss.getRange('Data!S3:S14');
    source.copyTo(ss.getRange('Week3!AA10:AA21'), {contentsOnly: true});
       var source = ss.getRange('Data!S16:S69');
    source.copyTo(ss.getRange('Week3!AA22:AA75'), {contentsOnly: true});
       var source = ss.getRange('Data!S71:S78');
    source.copyTo(ss.getRange('Week3!AA76:AA83'), {contentsOnly: true});
       var source = ss.getRange('Data!S82:S87');
    source.copyTo(ss.getRange('Week3!AA84:AA89'), {contentsOnly: true});
       var source = ss.getRange('Data!S89:S90');
    source.copyTo(ss.getRange('Week3!AA90:AA91'), {contentsOnly: true});
       var source = ss.getRange('Data!S92:S93');
    source.copyTo(ss.getRange('Week3!AA92:AA93'), {contentsOnly: true});
       var source = ss.getRange('Data!S95:S96');
    source.copyTo(ss.getRange('Week3!AA94:AA95'), {contentsOnly: true});
       var source = ss.getRange('Data!S98:S99');
    source.copyTo(ss.getRange('Week3!AA96:AA97'), {contentsOnly: true});
       var source = ss.getRange('Data!S101:S102');
    source.copyTo(ss.getRange('Week3!AA98:AA99'), {contentsOnly: true});

      //Column 6
            var source = ss.getRange('Data!T3:T14');
    source.copyTo(ss.getRange('Week3!AG10:AG21'), {contentsOnly: true});
       var source = ss.getRange('Data!T16:T69');
    source.copyTo(ss.getRange('Week3!AG22:AG75'), {contentsOnly: true});
       var source = ss.getRange('Data!T71:T78');
    source.copyTo(ss.getRange('Week3!AG76:AG83'), {contentsOnly: true});
       var source = ss.getRange('Data!T82:T87');
    source.copyTo(ss.getRange('Week3!AG84:AG89'), {contentsOnly: true});
       var source = ss.getRange('Data!T89:T90');
    source.copyTo(ss.getRange('Week3!AG90:AG91'), {contentsOnly: true});
       var source = ss.getRange('Data!T92:T93');
    source.copyTo(ss.getRange('Week3!AG92:AG93'), {contentsOnly: true});
       var source = ss.getRange('Data!T95:T96');
    source.copyTo(ss.getRange('Week3!AG94:AG95'), {contentsOnly: true});
       var source = ss.getRange('Data!T98:T99');
    source.copyTo(ss.getRange('Week3!AG96:AG97'), {contentsOnly: true});
       var source = ss.getRange('Data!T101:T102');
    source.copyTo(ss.getRange('Week3!AG98:AG99'), {contentsOnly: true});

      //Column 7
            var source = ss.getRange('Data!U3:U14');
    source.copyTo(ss.getRange('Week3!AM10:AM21'), {contentsOnly: true});
       var source = ss.getRange('Data!U16:U69');
    source.copyTo(ss.getRange('Week3!AM22:AM75'), {contentsOnly: true});
       var source = ss.getRange('Data!U71:U78');
    source.copyTo(ss.getRange('Week3!AM76:AM83'), {contentsOnly: true});
       var source = ss.getRange('Data!U82:U87');
    source.copyTo(ss.getRange('Week3!AM84:AM89'), {contentsOnly: true});
       var source = ss.getRange('Data!U89:U90');
    source.copyTo(ss.getRange('Week3!AM90:AM91'), {contentsOnly: true});
       var source = ss.getRange('Data!U92:U93');
    source.copyTo(ss.getRange('Week3!AM92:AM93'), {contentsOnly: true});
       var source = ss.getRange('Data!U95:U96');
    source.copyTo(ss.getRange('Week3!AM94:AM95'), {contentsOnly: true});
       var source = ss.getRange('Data!U98:U99');
    source.copyTo(ss.getRange('Week3!AM96:AM97'), {contentsOnly: true});
       var source = ss.getRange('Data!U101:U102');
    source.copyTo(ss.getRange('Week3!AM98:AM99'), {contentsOnly: true});

      //Column 8
      
            var source = ss.getRange('Data!V3:V14');
    source.copyTo(ss.getRange('Week3!AS10:AS21'), {contentsOnly: true});
       var source = ss.getRange('Data!V16:V69');
    source.copyTo(ss.getRange('Week3!AS22:AS75'), {contentsOnly: true});
       var source = ss.getRange('Data!V71:V78');
    source.copyTo(ss.getRange('Week3!AS76:AS83'), {contentsOnly: true});
       var source = ss.getRange('Data!V82:V87');
    source.copyTo(ss.getRange('Week3!AS84:AS89'), {contentsOnly: true});
       var source = ss.getRange('Data!V89:V90');
    source.copyTo(ss.getRange('Week3!AS90:AS91'), {contentsOnly: true});
       var source = ss.getRange('Data!V92:V93');
    source.copyTo(ss.getRange('Week3!AS92:AS93'), {contentsOnly: true});
       var source = ss.getRange('Data!V95:V96');
    source.copyTo(ss.getRange('Week3!AS94:AS95'), {contentsOnly: true});
       var source = ss.getRange('Data!V98:V99');
    source.copyTo(ss.getRange('Week3!AS96:AS97'), {contentsOnly: true});
       var source = ss.getRange('Data!V101:V102');
    source.copyTo(ss.getRange('Week3!AS98:AS99'), {contentsOnly: true});

      
      //Molecules
      
      //Column 1
      var source = ss.getRange('Molecule!O3:O14');
    source.copyTo(ss.getRange('Week3!B10:B21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!O16:O69');
    source.copyTo(ss.getRange('Week3!B22:B75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!O71:O78');
    source.copyTo(ss.getRange('Week3!B76:B83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!O82:O87');
    source.copyTo(ss.getRange('Week3!B84:B89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!O89:O90');
    source.copyTo(ss.getRange('Week3!B90:B91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!O92:O93');
    source.copyTo(ss.getRange('Week3!B92:B93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!O95:O96');
    source.copyTo(ss.getRange('Week3!B94:B95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!O98:O99');
    source.copyTo(ss.getRange('Week3!B96:B97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!O101:O102');
    source.copyTo(ss.getRange('Week3!B98:B99'), {contentsOnly: true});

      //Column 2
      var source = ss.getRange('Molecule!P3:P14');
    source.copyTo(ss.getRange('Week3!H10:H21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!P16:P69');
    source.copyTo(ss.getRange('Week3!H22:H75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!P71:P78');
    source.copyTo(ss.getRange('Week3!H76:H83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!P82:P87');
    source.copyTo(ss.getRange('Week3!H84:H89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!P89:P90');
    source.copyTo(ss.getRange('Week3!H90:H91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!P92:P93');
    source.copyTo(ss.getRange('Week3!H92:H93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!P95:P96');
    source.copyTo(ss.getRange('Week3!H94:H95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!P98:P99');
    source.copyTo(ss.getRange('Week3!H96:H97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!P101:P102');
    source.copyTo(ss.getRange('Week3!H98:H99'), {contentsOnly: true});

      //Column 3
      var source = ss.getRange('Molecule!Q3:Q14');
    source.copyTo(ss.getRange('Week3!N10:N21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!Q16:Q69');
    source.copyTo(ss.getRange('Week3!N22:N75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!Q71:Q78');
    source.copyTo(ss.getRange('Week3!N76:N83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!Q82:Q87');
    source.copyTo(ss.getRange('Week3!N84:N89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!Q89:Q90');
    source.copyTo(ss.getRange('Week3!N90:N91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!Q92:Q93');
    source.copyTo(ss.getRange('Week3!N92:N93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!Q95:Q96');
    source.copyTo(ss.getRange('Week3!N94:N95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!Q98:Q99');
    source.copyTo(ss.getRange('Week3!N96:N97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!Q101:Q102');
    source.copyTo(ss.getRange('Week3!N98:N99'), {contentsOnly: true});

      //Column 4
      var source = ss.getRange('Molecule!R3:R14');
    source.copyTo(ss.getRange('Week3!T10:T21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!R16:R69');
    source.copyTo(ss.getRange('Week3!T22:T75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!R71:R78');
    source.copyTo(ss.getRange('Week3!T76:T83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!R82:R87');
    source.copyTo(ss.getRange('Week3!T84:T89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!R89:R90');
    source.copyTo(ss.getRange('Week3!T90:T91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!R92:R93');
    source.copyTo(ss.getRange('Week3!T92:T93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!R95:R96');
    source.copyTo(ss.getRange('Week3!T94:T95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!R98:R99');
    source.copyTo(ss.getRange('Week3!T96:T97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!R101:R102');
    source.copyTo(ss.getRange('Week3!T98:T99'), {contentsOnly: true});

      //Column 5
      var source = ss.getRange('Molecule!S3:S14');
    source.copyTo(ss.getRange('Week3!Z10:Z21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!S16:S69');
    source.copyTo(ss.getRange('Week3!Z22:Z75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!S71:S78');
    source.copyTo(ss.getRange('Week3!Z76:Z83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!S82:S87');
    source.copyTo(ss.getRange('Week3!Z84:Z89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!S89:S90');
    source.copyTo(ss.getRange('Week3!Z90:Z91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!S92:S93');
    source.copyTo(ss.getRange('Week3!Z92:Z93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!S95:S96');
    source.copyTo(ss.getRange('Week3!Z94:Z95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!S98:S99');
    source.copyTo(ss.getRange('Week3!Z96:Z97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!S101:S102');
    source.copyTo(ss.getRange('Week3!Z98:Z99'), {contentsOnly: true});

      //Column 6
      var source = ss.getRange('Molecule!T3:T14');
    source.copyTo(ss.getRange('Week3!AF10:AF21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!T16:T69');
    source.copyTo(ss.getRange('Week3!AF22:AF75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!T71:T78');
    source.copyTo(ss.getRange('Week3!AF76:AF83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!T82:T87');
    source.copyTo(ss.getRange('Week3!AF84:AF89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!T89:T90');
    source.copyTo(ss.getRange('Week3!AF90:AF91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!T92:T93');
    source.copyTo(ss.getRange('Week3!AF92:AF93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!T95:T96');
    source.copyTo(ss.getRange('Week3!AF94:AF95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!T98:T99');
    source.copyTo(ss.getRange('Week3!AF96:AF97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!T101:T102');
    source.copyTo(ss.getRange('Week3!AF98:AF99'), {contentsOnly: true});

      //Column 7
      var source = ss.getRange('Molecule!U3:U14');
    source.copyTo(ss.getRange('Week3!AL10:AL21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!U16:U69');
    source.copyTo(ss.getRange('Week3!AL22:AL75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!U71:U78');
    source.copyTo(ss.getRange('Week3!AL76:AL83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!U82:U87');
    source.copyTo(ss.getRange('Week3!AL84:AL89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!U89:U90');
    source.copyTo(ss.getRange('Week3!AL90:AL91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!U92:U93');
    source.copyTo(ss.getRange('Week3!AL92:AL93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!U95:U96');
    source.copyTo(ss.getRange('Week3!AL94:AL95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!U98:U99');
    source.copyTo(ss.getRange('Week3!AL96:AL97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!U101:U102');
    source.copyTo(ss.getRange('Week3!AL98:AL99'), {contentsOnly: true});

      //Column 8
      var source = ss.getRange('Molecule!V3:V14');
    source.copyTo(ss.getRange('Week3!AR10:AR21'), {contentsOnly: true});
       var source = ss.getRange('Molecule!V16:V69');
    source.copyTo(ss.getRange('Week3!AR22:AR75'), {contentsOnly: true});
       var source = ss.getRange('Molecule!V71:V78');
    source.copyTo(ss.getRange('Week3!AR76:AR83'), {contentsOnly: true});
       var source = ss.getRange('Molecule!V82:V87');
    source.copyTo(ss.getRange('Week3!AR84:AR89'), {contentsOnly: true});
       var source = ss.getRange('Molecule!V89:V90');
    source.copyTo(ss.getRange('Week3!AR90:AR91'), {contentsOnly: true});
       var source = ss.getRange('Molecule!V92:V93');
    source.copyTo(ss.getRange('Week3!AR92:AR93'), {contentsOnly: true});
       var source = ss.getRange('Molecule!V95:V96');
    source.copyTo(ss.getRange('Week3!AR94:AR95'), {contentsOnly: true});
       var source = ss.getRange('Molecule!V98:V99');
    source.copyTo(ss.getRange('Week3!AR96:AR97'), {contentsOnly: true});
       var source = ss.getRange('Molecule!V101:V102');
    source.copyTo(ss.getRange('Week3!AR98:AR99'), {contentsOnly: true});

      
    }
*/

  


