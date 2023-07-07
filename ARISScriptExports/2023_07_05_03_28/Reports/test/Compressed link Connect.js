function compressGuid(guid) {
    var code = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_";
    var compressedGuid = "";
    var bucket = 0;
    var bits = 0;
    //for (var i = 0; i < guid.length(); i++) {
    for (var i = 0; i < guid.length; i++) {
 
        //var rValue = guid.charAt(i)
        var rValue = guid.charCodeAt(i);
        if (rValue >= 48 && rValue <=57)
            fValue = rValue - 48;
        else if (rValue >= 97 && rValue <= 102)
            fValue = 10 + rValue - 97;
        else if (rValue >= 65 && rValue <= 90)
            fValue = 10 + rValue - 65;
        else            
            continue;
 
       bucket = (bucket << 4) + fValue;
        bits+=4;
        if (bits >= 6) {
            bits-=6;
            sValue = bucket >> bits;
            bucket = bucket ^ (sValue << bits);
            compressedGuid += code.charAt(sValue);
        }
    }
    if (bits > 0) {
        compressedGuid += code.charAt(bucket << (6 - bits));
    }
    return compressedGuid;
}

function getLink(sGUID, sFactsheet, iVersion) {

    var CONNECT_URL = "http://lubuntu-aris:1080";
    var g_nloc = Context.getSelectedLanguage();
    
    if (iVersion == null) iVersion = -1;    
    (sFactsheet == null) ? sFactsheet = "" : sFactsheet = sFactsheet + ".";
    
    var oDatabase = ArisData.getActiveDatabase();
    var sTenantName = ArisData.getTenantName();
    var serverName = oDatabase.ServerName ( );
    var sDBName = oDatabase.Name(g_nloc);

    sDBName = sDBName.replace(".", "~d");
    
    var linkURL = CONNECT_URL + "/#" + sTenantName + "/item/c." + sFactsheet + sDBName + "." + compressGuid(sGUID) + "." + iVersion;
    
    var encodedURL = encodeURI(linkURL);
    
    return encodedURL;
    
    
}

getLink("bc1b3230-f98c-11ed-2ba6-0a002700001a", null, null)
// returns:  http://lubuntu-aris:1080/#default/item/c.ARISApp%20-%20ARIS%20Notifier%20-%20Filter.xZCl4FUIEecMzQAcQq7k8Q.-1