function Product_Pricing()
{
  CreateKFI('product_maintenance');
}

function Product_Special_Pricing()
{
  CreateKFI('product_special_pricing');
}

function CreateKFI(kfiType)
{
  // Get the spreadsheet object so we can read from it
  var Spreadsheet = SpreadsheetApp.getActiveSheet();
  var productDataRange = Spreadsheet.getDataRange();
  var productRange = Spreadsheet.getRange(1, 1, productDataRange.getLastRow(), productDataRange.getLastColumn());
  var productValues = productRange.getValues();
  
  // Row number for where we need to look for product listings (starts at 0 not 1)
  var firstProductRow = 12;
  
  // Get the list of products from the spreadsheet by calling a function defined below
  var products = getProductListing();
  
  // Setup the KFI file we're creating
  var customFileName = productValues[firstProductRow - 2][2].toString().trim().replace(/\.| |\/|\\/, '')
  var fileName = customFileName != '' ? customFileName : (new Date()).toString();
  var fileStart, fileEnd, fileBody;
  
  if (kfiType == 'product_maintenance')
  {
    fileStart = '<?xml version="1.0" encoding="UTF-8"?><xmlkfi object_id="19">\r\n';
    fileEnd = '</xmlkfi>';
  
    // Create the text string based on the contents of the products array of product objects
    fileBody = products.map(function(a) {
      if (a.model == '')
        return '';
      
      return '\t<record>\r\n\t\t<location></location>\r\n\t\t<code>' + a.model + '</code>\r\n\t\t<desc></desc>\r\n\t\t<inactive></inactive>\r\n'
      + (a.price ? ( '\t\t<sell1>' + a.price + '</sell1>\r\n' ) : '' )
      + (a.webSpecialPrice != null ? ('\t\t<decmlnum01>' + a.webSpecialPrice + '</decmlnum01>\r\n') : '')
      + (a.webSpecialExpiry != null ? ('\t\t<date01>' + getIncrementedExpiry( a.webSpecialExpiry, 1 ) + '</date01>\r\n') : '')
      + (a.webSpecialFlag != null ? '\t\t<yesno03>' + a.webSpecialFlag + '</yesno03>\r\n' : '')
      + "\t</record>\r\n";
    }).join('');
  }
  else if (kfiType == 'product_special_pricing') {
    fileStart = '<?xml version="1.0" encoding="UTF-8"?><xmlkfi object_id="133">';
    fileEnd = '</xmlkfi>';
  
    // Create the text string based on the contents of the products array of product objects
    fileBody = products.map(function (a) {
      if (a.model == '' || a.webSpecialStart == null || a.webSpecialExpiry == null || a.webSpecialPrice == -1)
      {
        throw 'Error in generation of Customer_Special_Pricing KFI - incorrect parameters specified ...';
        return '';
      }
      
      var isDates = a.webSpecialStart != null && a.webSpecialExpiry != null;
      
      return '\r\n\t<record>\r\n\t\t<header>\r\n\t\t\t<cuscode></cuscode>\r\n\t\t\t<prodcode>' + a.model + '</prodcode>\r\n\t\t\t<cusprodfld></cusprodfld>'
        + ( isDates ? ( '\r\n\t\t\t<daterange>Y</daterange>\r\n\t\t\t\t<valfrom>' + a.webSpecialStart + '</valfrom>\r\n\t\t\t\t<valto>' + a.webSpecialExpiry + '</valto>' ) : '\r\n\t\t\t<daterange>N</daterange>' )
        + '\r\n\t\t</header>\r\n\t\t<pricelevel>\r\n\t\t\t<maxqty>9999999999</maxqty>\r\n\t\t\t<prccde>9</prccde>\r\n\t\t\t<unitpr>' + a.webSpecialPrice + '</unitpr>\r\n\t\t</pricelevel>\r\n\t</record>';
    } ).join( '' );
  }
  
  DriveApp.getFolderById('1j4ihsEtTzSgxaNqG9pr1fauPNaFj0rmM').createFile(fileName + '_' + kfiType + '.kfi', fileStart + fileBody + fileEnd, MimeType.PLAIN_TEXT);
  
  function getProductListing()
  {
    var productList = [];
    var clearCommand = '&lt;BKSP&gt;';
    
    for ( var i = firstProductRow; i < productValues.length; i++ )
    {
      var model = productValues[ i ][ 0 ].toString().trim();
      var price = parseFloat(productValues[i][1]) || -1;
      var webSpecialPrice = parseFloat(productValues[i][2]) || -1;
      var webSpecialFlag = productValues[i][3] == 'Y' ? 'Y' : (productValues[i][3].trim() == '' ? '' : 'N');
      var webSpecialStart = productValues[i][4].toString().replace(/(20)(?=[0-9]{2})/g, '').replace(/\/|\.|\\| /g, '').trim();
      var webSpecialExpiry = productValues[i][5].toString().replace(/(20)(?=[0-9]{2})/g, '').replace(/\/|\.|\\| /g, '').trim();
      
      var obj = { model: model };
      
      if (price != -1)
        obj.price = price;
      
      if (webSpecialPrice != -1)
        obj.webSpecialPrice = webSpecialPrice;
      
      if (webSpecialFlag != '')
        obj.webSpecialFlag = webSpecialFlag;
      
      if (webSpecialStart != '')
        obj.webSpecialStart = webSpecialStart.length == 5 ? '0' + webSpecialStart : webSpecialStart;
      
      if (webSpecialExpiry != '')
        obj.webSpecialExpiry = webSpecialExpiry.length == 5 ? '0' + webSpecialExpiry : webSpecialExpiry;
      
      Logger.log(webSpecialStart + '    ' + webSpecialStart.length);
      
      // Only add to the product list array if we're changing at least one property
      if (Object.keys(obj).length > 1)
        productList.push(obj);
    }
    
    return productList;
  }
  
  function getIncrementedExpiry(_date, _amount)
  {
    var _d = '20' + _date.substring(4) + '-' + _date.substring(2, 4) + '-' + _date.substring(0, 2);
    var increment = (new Date(new Date(_d).setDate(new Date(_d).getDate() + _amount))).toISOString().substring(2, 10);
    return increment.substring(6) + increment.substring(3, 5) + increment.substring(0, 2);
  }
  
  function isCurrentSaleDate(_dateString)
  {
    var dObj = new Date(_dateString.substring(4, 8) + '-' + _dateString.substring(2, 4) + '-' + _dateString.substring(0, 2) + 'T00:00:00');
    return dObj.getTime() - (new Date()).getTime() > 0;
  }
}
