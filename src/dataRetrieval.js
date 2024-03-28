
// last synchronized Oct 13 2018 Version 1.0

// adapted from: https://www.quora.com/How-can-I-exceed-the-50-importxml-limit-in-Google-Sheets

/**
 * dataRetrieval: attempts to retrieve a CAS number from Sigma Aldrich 
 *
 * @param {string} Compoundname The name or abbreviation of the compound to look for.
 * @return The CAS number found on the Aldrich website.
 * @customfunction
 */

function CASfromAldrich(Compoundname) {
  var output = '';
  //Compoundname = "Na2S"
  Compoundname = Compoundname.replace(" ", "+");
  Compoundname = Compoundname.replace("%", "");
  var url = "https://www.sigmaaldrich.com/catalog/search?term=" + Compoundname + "&interface=All&N=0&mode=match%20partialmax&lang=en&region=US&focus=product";
  var fetchedUrl = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (fetchedUrl) {
    var html = fetchedUrl.getContentText();
    if (html.length) {
      output = html.match(new RegExp("[0-9]{2,7}-[0-9]{2}-[0-9]", 'i'));
    }
  }
  // Grace period to avoid call limit
  Utilities.sleep(500);
  //Logger.log(output)
  if (output == null)
    throw 'nothing found at Aldrich';
  else
    return output;

}

/**
 * dataRetrieval: attempts to retrieve a CAS number from the ChemicalBook website
 *
 * @param {string} input The name or abbreviation of the compound to look for.
 * @return The CAS number found on the ChemicalBook website.
 * @customfunction
 */

function CASfromChemicalBook(Compoundname) {
  var output = '';

  Compoundname = "Na2CO3";
  Compoundname = Compoundname.replace("%", "");
  Compoundname = Compoundname.replace(" ", "%20");

  var url = "http://www.chemicalbook.com/ProductList_En.aspx?kwd=" + Compoundname;
  var fetchedUrl = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (fetchedUrl) {
    var html = fetchedUrl.getContentText();
    Logger.log(html);
    if (html.length) {
      output = RegExExtractAllMatches(html, "[0-9]{2,7}-[0-9]{2}-[0-9]", "i");
      output = output[0];
      output = output.slice(2, output.length);

      if (output.length > 1) { output = mostcommonelement(output); }
    }
  }
  // Grace period to avoid call limit
  Utilities.sleep(500);
  if (output.length == 0)
    throw 'nothing found at Chemical Book';
  else
    return output;
}

/**
 * dataRetrieval: attempts to retrieve a CAS number from Google
 *
 * @param {string} input The name or abbreviation of the compound to look for.
 * @return The CAS number found on Google.
 * @customfunction
 */

function CASfromGoogle(Compoundname) {
  var output = '';

  Compoundname = "Na2CO3";
  Compoundname = Compoundname.replace("%", "");
  Compoundname = Compoundname.replace(" ", "+");

  var url = "https://www.google.com/search?q=" + Compoundname + "+CAS";
  var fetchedUrl = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (fetchedUrl) {
    var html = fetchedUrl.getContentText();

    if (html.length) {
      output = RegExExtractAllMatches(html, "[0-9]{2,7}-[0-9]{2}-[0-9]", "i");

      output = output[0];

      if (output == 0)
        throw 'nothing found';
      if (output.length > 1) {
        //from all CAS-numbers found, choose the one found most often
        output = mostcommonelement(output);
        //checks whether CAS-number is valid and throws an error if not
        if (validateCAS(output) == 0)
          throw "not a valid CAS-number";

      }

    }
  }
  // Grace period to avoid call limit
  Utilities.sleep(1000);
  if (validateCAS(output) == 0)
    throw "not a valid CAS-number";
  return output;
}

// This work is hereby released into the public domain. To the extent
// possible under law, the author has waived all copyright and related
// or neighboring rights to the work. See the dedication at
//
//     http://creativecommons.org/publicdomain/zero/1.0/
//
// The author makes no warranties about the work, and disclaims liability
// for all uses of the work, to the fullest extent permitted by applicable
// law. When using or citing the work, you should not imply endorsement by
// the author.
/**
* dataRetrieval: Returns all occurrences of a text string in a cell.
*
* Accepts range parameters and matches the occurrences across each row separately.
* The result will have as many rows as the parameter range and as many columns as there are results.
* Note that matching a special character requires \escaping as in \* \+ \. \$ \[ \(.
*
* @param {A2:D42} text_to_search The text string, cell or range to search.
* @param {"abc|def"} search_for The text string or regular expression to search for within text_to_search.
* @param {true} ignore_case Set true to treat lowercase letters the same as uppercase.
* @return The occurrences of a regular expressions in a cell. When the argument is an array, returns an array of matches.
* @customfunction
*/
function RegExExtractAllMatches(text_to_search, search_for, ignore_case) {
  // version 1.0, written by --Hyde, 25 September 2014
  if (!search_for)
    throw 'Empty search string given to RegExExtractAllMatches.';
  if (arguments.length < 2 || arguments.length > 3)
    throw 'Wrong number of arguments to RegExExtractAllMatches. Expected 2 or 3 arguments, but got ' + arguments.length + ' arguments.';

  if (text_to_search.constructor != Array) text_to_search = [[text_to_search]];
  search_for = new RegExp(search_for, "g" + (ignore_case ? "i" : ""));
  var result = [];

  for (var i = 0; i < text_to_search.length; i++) {
    result[i] = text_to_search[i].join(String.fromCharCode(9)).match(search_for);
  }
  return result;
}

/**
 * dataRetrieval: get the element that's found most often in a given array. Adapted from: https://stackoverflow.com/questions/1053843/get-the-element-with-the-highest-occurrence-in-an-array
 * @param {Array} array An array containing duplicates.
 * @return {String} maxEl.
 */
function mostcommonelement(array) {
  if (array.length == 0)
    return null;
  var modeMap = {};
  var maxEl = array[0], maxCount = 1;
  for (var i = 0; i < array.length; i++) {
    var el = array[i];
    if (modeMap[el] == null)
      modeMap[el] = 1;
    else
      modeMap[el]++;
    if (modeMap[el] > maxCount) {
      maxEl = el;
      maxCount = modeMap[el];
    }
  }
  return maxEl;
}

/**
 * dataRetrieval: Use Google's customsearch API to perform a search query.
 * See https://developers.google.com/custom-search/json-api/v1/using_rest.
 *
 * @param {string} query   Search query to perform, e.g. "test"
 * @return The CAS number found on Google.
 * @customfunction
 * returns {object}        See response data structure at
 *                         https://developers.google.com/custom-search/json-api/v1/reference/cse/list#response
 */
function searchForCAS(query) {

  // Base URL to access customsearch
  var urlTemplate = "https://www.googleapis.com/customsearch/v1?key=%KEY%&cx=%CX%&q=%Q%";
  //query = "NaOH CAS"
  // Script-specific credentials & search engine


  // Build custom url
  var url = urlTemplate
    .replace("%KEY%", encodeURIComponent(ApiKey))
    .replace("%CX%", encodeURIComponent(SearchEngineID))
    .replace("%Q%", encodeURIComponent(query));

  var params = {
    muteHttpExceptions: true
  };

  // Perform search
  Logger.log(UrlFetchApp.getRequest(url, params));  // Log query to be sent
  var response = UrlFetchApp.fetch(url, params);
  var respCode = response.getResponseCode();

  if (respCode !== 200) {
    throw new Error("Error " + respCode + " " + response.getContentText());
  }
  else {
    // Successful search, log & return results
    var output = response.getContentText();
    output = RegExExtractAllMatches(output, "[0-9]{2,7}-[0-9]{2}-[0-9]", "i");
    Logger.log(output);
    output = mostcommonelement(output[0]);
    Logger.log(output);
    return output;
  }
}

//from: http://depth-first.com/articles/2011/10/20/how-to-validate-cas-registry-numbers-in-javascript/

/**
 * dataRetrieval: check whether a given string is a valid CAS-number
 * @param {string} CAS A candidate to be checked whether it's a valid CAS number.
 * @return {Number} the checksum result indicating whether it's a valid CAS number.
 */
function validateCAS(CAS) {
  var sum = 0;

  Logger.log(CAS);
  var digits = CAS.replace("-", '');
  digits = digits.replace("-", '');

  for (var i = digits.length - 2; i >= 0; i--) {
    sum = sum + Number(digits[i]) * (digits.length - i - 1);
  }

  return sum % 10 === Number(CAS.slice(-1));
}