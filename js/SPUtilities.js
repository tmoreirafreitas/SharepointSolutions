// De https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object/keys

if (!Object.keys) {
  Object.keys = (function() {
    'use strict';
    var hasOwnProperty = Object.prototype.hasOwnProperty,
        hasDontEnumBug = !({ toString: null }).propertyIsEnumerable('toString'),
        dontEnums = [
          'toString',
          'toLocaleString',
          'valueOf',
          'hasOwnProperty',
          'isPrototypeOf',
          'propertyIsEnumerable',
          'constructor'
        ],
        dontEnumsLength = dontEnums.length;

    return function(obj) {
      if (typeof obj !== 'object' && (typeof obj !== 'function' || obj === null)) {
        throw new TypeError('Object.keys chamado de non-object');
      }

      var result = [], prop, i;

      for (prop in obj) {
        if (hasOwnProperty.call(obj, prop)) {
          result.push(prop);
        }
      }

      if (hasDontEnumBug) {
        for (i = 0; i < dontEnumsLength; i++) {
          if (hasOwnProperty.call(obj, dontEnums[i])) {
            result.push(dontEnums[i]);
          }
        }
      }
      return result;
    };
  }());
}

(function() {

  function DateDiff(date1, date2) {
	this.years = null;
	this.months = null;
	this.weeks = null;
    this.days = null;
    this.hours = null;
    this.minutes = null;
    this.seconds = null;	
    this.date1 = date1;
    this.date2 = date2;

    this.init();
  }

  DateDiff.prototype.init = function() {
    var data = new DateMeasure(this.date1 - this.date2);
	this.years = data.years;
	this.months = data.months;
	this.weeks = data.weeks;
    this.days = data.days;
    this.hours = data.hours;
    this.minutes = data.minutes;
    this.seconds = data.seconds;	
  };

  function DateMeasure(ms) {
    var y,mth,w,d, h, m, s;    	
	y = Math.floor(ms / 31536000000);
	mth = Math.floor(ms / (60*60*24*7*4*1000));
	w = Math.floor(ms / 604800000);
	d = Math.floor(ms / 86400000);
	h = Math.floor(ms / 3600000);
	m = Math.floor(ms / 60000);
	s = Math.floor(ms / 1000);
    
	this.years = y;
	this.months = mth;
	this.weeks = w;
    this.days = d;
    this.hours = h;
    this.minutes = m;
    this.seconds = s;
  };

  Date.diff = function(date1, date2) {
    return new DateDiff(date1, date2);
  };

  Date.prototype.diff = function(date2) {
    return new DateDiff(this, date2);
  };

})();

if (!Array.prototype.indexOf)
{
  Array.prototype.indexOf = function(elt /*, from*/)
  {
    var len = this.length >>> 0;

    var from = Number(arguments[1]) || 0;
    from = (from < 0)
         ? Math.ceil(from)
         : Math.floor(from);
    if (from < 0)
      from += len;

    for (; from < len; from++)
    {
      if (from in this &&
          this[from] === elt)
        return from;
    }
    return -1;
  };
}

if (!Date.prototype.toISOString) {
  (function() {

    function pad(number) {
      if (number < 10) {
        return '0' + number;
      }
      return number;
    }

    Date.prototype.toISOString = function() {
      return this.getUTCFullYear() +
        '-' + pad(this.getUTCMonth() + 1) +
        '-' + pad(this.getUTCDate()) +
        'T' + pad(this.getUTCHours()) +
        ':' + pad(this.getUTCMinutes()) +
        ':' + pad(this.getUTCSeconds()) +
        '.' + (this.getUTCMilliseconds() / 1000).toFixed(3).slice(2, 5) +
        'Z';
    };

  }());
}

String.prototype.fromEdmDate = function() {
	return new Date(parseInt(this.match(/\/Date\(([0-9]+)(?:.*)\)\//)[1]));
};


function isNumberKey(evt){
	var charCode = (evt.which) ? evt.which : event.keyCode
	if (charCode > 31 && (charCode < 48 || charCode > 57))
		return false;

	return true;
}

var defineProp = function ( obj, key, value ){
  var config = {
    value: value,
    writable: true,
    enumerable: true,
    configurable: true
  };
  Object.defineProperty( obj, key, config );
};

function getAttributes ( $node ) {
    var attrs = {};
    $.each( $node[0].attributes, function ( index, attribute ) {
        attrs[attribute.name] = attribute.value;
    } );

    return attrs;
}

var groupBy = function(xs, key) {
  return xs.reduce(function(rv, x) {
		(rv[x[key].toLowerCase()] = rv[x[key].toLowerCase()] || []).push(x);
		return rv;
  }, {});
};

function removeDuplicates(originalArray, property) {
     var newArray = [];
     var lookupObject  = {};
     for(var i in originalArray) {
        lookupObject[originalArray[i][property]] = originalArray[i];
     }
     for(i in lookupObject) {
         newArray.push(lookupObject[i]);
     }
      return newArray;
 }

 function getParameterByName(name, url) {
    if (!url) url = window.location.href;
    name = name.replace(/[\[\]]/g, "\\$&");
    var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
        results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, " "));
}

function convertThisDate(thisDate){
  var stringDate = "";
  stringDate  += thisDate.getFullYear() + "-";
  stringDate  += (((thisDate.getMonth() + 1).toString().length) == 1 ? ("0" + (thisDate.getMonth() + 1)) : (thisDate.getMonth() + 1)) + "-";
  stringDate  += thisDate.getDate().toString().length == 1 ? ("0" + thisDate.getDate()) : thisDate.getDate();
  stringDate  += "T" + (thisDate.getHours().toString().length == 1 ? ("0" + thisDate.getHours()) : thisDate.getHours()) +
   ":" + (thisDate.getMinutes().toString().length == 1 ? ("0" + thisDate.getMinutes()) : thisDate.getMinutes()) + 
   ":" + (thisDate.getSeconds().toString().length == 1 ? ("0" + thisDate.getSeconds()) : thisDate.getSeconds());

  return stringDate;
}

function hasWhiteSpace(s) {
  return /\s/g.test(s);
}

function verificaNumero(e) {
    if (e.which != 8 && e.which != 0 && (e.which < 48 || e.which > 57)) {
        return false;
    }
}

function getCurrentUserName(){
	var currentName = $().SPServices.SPGetCurrentUser({
		async: false,
  		webURL: "",
  		fieldName: "Title",
  		fieldNames: {},
  		debug: false
	});

	return currentName;
}

function getCurrentUserLogin(){
	var currentLogin = $().SPServices.SPGetCurrentUser({
		async: false,
  		webURL: "",
  		fieldName: "Name",
  		fieldNames: {},
  		debug: false
	});

	return currentLogin;
}

function createListItem(fieldValues, listName){
    var def = new $.Deferred();
    var valuePair = fieldValues;
	var collListItem = [];

    //SPServices to Create an Item.	
	$().SPServices({
		operation: "UpdateListItems",
		async: false,
		webURL: $().SPServices.SPGetCurrentSite(),
		listName: listName,
		batchCmd: "New",
		valuepairs: valuePair,
		completefunc: function (xData, Status) {
			if (Status == "success") {
				/*var newId = $(xData.responseXML).SPFilterNode("z:row").attr("ows_ID");
				newItemID = newId;
				def.resolve(newItemID);*/
				var collObj = [];				
				$(xData.responseXML).SPFilterNode("z:row").each(function() {
					var obj = getAttributes($(this));
					collObj.push(obj);
				});
				collListItem = collObj;
				def.resolve(collListItem);
			}
			else{
				var spErrorcode =  $(xData.responseXML).SPFilterNode("ErrorCode").text(); // 0x00000000 
				var spErrortext =  $(xData.responseXML).SPFilterNode("ErrorText").text() + $(xData.responseXML).SPFilterNode("errorstring").text();
				def.reject('Error in SPServices call\n SharePoint errorcode=' + spErrorcode + '\n ErrorText=' + spErrortext + '\n SPServices status=' + Status);
			}
		}
	});
    return def;
}

function updateListItem(id, valuePair, listName){
    var def = new $.Deferred();    
    var itemID = id;
	var updateValuePair = valuePair;
	var collListItem = [];
	
    //SPServices to Create an Item.	
	$().SPServices({
		operation: "UpdateListItems",
		async: false,		
		webURL: $().SPServices.SPGetCurrentSite(),
		listName: listName,
		batchCmd: "Update",
        ID: itemID,
		valuepairs: updateValuePair,
		completefunc: function (xData, Status) {
			if (Status == "success") {				
				var collObj = [];				
				$(xData.responseXML).SPFilterNode("z:row").each(function() {
					var obj = getAttributes($(this));
					collObj.push(obj);
				});
				collListItem = collObj;
				def.resolve(collListItem);
			}
			else {
				var spErrorcode =  $(xData.responseXML).SPFilterNode("ErrorCode").text(); // 0x00000000 
				var spErrortext =  $(xData.responseXML).SPFilterNode("ErrorText").text() + $(xData.responseXML).SPFilterNode("errorstring").text();
				def.reject('Error in SPServices call\n SharePoint errorcode=' + spErrorcode + '\n ErrorText=' + spErrortext + '\n SPServices status=' + Status);
			}
		}
	});
    return def;
}

function deleteListItem(itemId, listName){
    var def = new $.Deferred();
    $().SPServices(  
    {  
        operation: "UpdateListItems",  
        async: false,  
        webURL: $().SPServices.SPGetCurrentSite(),
        batchCmd: "Delete",  
        listName: listName,  
        ID: itemId,  
        completefunc: function(xData, Status) {
			if (Status == "success") {					
				def.resolve("Item Id:" + itemId + " Deleted");
			}
			else {
				var spErrorcode =  $(xData.responseXML).SPFilterNode("ErrorCode").text(); // 0x00000000 
				var spErrortext =  $(xData.responseXML).SPFilterNode("ErrorText").text() + $(xData.responseXML).SPFilterNode("errorstring").text();
				def.reject('Error in SPServices call\n SharePoint errorcode=' + spErrorcode + '\n ErrorText=' + spErrortext + '\n SPServices status=' + Status);
			}            
        }  
    });  

    return def;
}

function getListItemByLookupId(list, fieldLookup, itemId){
	var def = new $.Deferred();				
	var collListItem = [];
	var query = "<Query>"+
					"<Where>"+
						"<Eq><FieldRef Name='" + fieldLookup + "' LookupId='True' /><Value Type='Lookup'>" + itemId + "</Value></Eq>"+
					"</Where>"+
				"</Query>";	
	$().SPServices({
	  webURL:$().SPServices.SPGetCurrentSite(),
      operation: 'GetListItems',
      listName: list,
	  CAMLQuery: query,
	  async: false,
      completefunc: function (xData, Status) {
			if (Status == "success") {
				var collObj = [];				
				$(xData.responseXML).SPFilterNode("z:row").each(function() {
					var obj = getAttributes($(this));
					collObj.push(obj);
				});
				collListItem = collObj;
				def.resolve(collListItem);
			}
			else {
				var spErrorcode =  $(xData.responseXML).find("ErrorCode").text(); // 0x00000000 
				var spErrortext =  $(xData.responseXML).find("ErrorText").text() + $(xData.responseXML).find("errorstring").text();
				def.reject('Error in SPServices call\n SharePoint errorcode=' + spErrorcode + '\n ErrorText=' + spErrortext + '\n SPServices status=' + Status);
			}
		}
    });
	return def;
}

function getListItemByChoice(list, fieldChoice, choiceValue){
	var def = new $.Deferred();    
	var listItem = null;
	var collListItem = [];
	var query = "<Query>"+
					 "<Where>"+
    					"<Eq>"+
         					"<FieldRef Name='" + fieldChoice + "' />"+
         					"<Value Type='Choice'>" + choiceValue + "</Value>"+
      					"</Eq>"+
   					"</Where>"+
				"</Query>";

	$().SPServices({
	  webURL: $().SPServices.SPGetCurrentSite(),
      operation: 'GetListItems',
      listName: list,
	  CAMLQuery: query,
	  async: false,
      completefunc: function (xData, Status) {
			if (Status == "success") {
				var collObj = [];
				$(xData.responseXML).SPFilterNode("z:row").each(function() {
					var obj = getAttributes($(this));
					collObj.push(obj);
				});
				collListItem = collObj;
				def.resolve(collListItem);
			}
			else {
				var spErrorcode =  $(xData.responseXML).find("ErrorCode").text(); // 0x00000000 
				var spErrortext =  $(xData.responseXML).find("ErrorText").text() + $(xData.responseXML).find("errorstring").text();
				def.reject('Error in SPServices call\n SharePoint errorcode=' + spErrorcode + '\n ErrorText=' + spErrortext + '\n SPServices status=' + Status);
			}
		}
    });
	return def;
}

function getListItemById(list, itemId){
    var def = new $.Deferred();
	var collListItem = [];
	var query =  "<Query>"+
					"<Where>"+
						"<Eq><FieldRef Name='ID' /><Value Type='Counter'>" + itemId + "</Value></Eq>"+
					"</Where>"+
				"</Query>";
															
	$().SPServices({
	  webURL: $().SPServices.SPGetCurrentSite(),
      operation: 'GetListItems',
      listName: list,
	  CAMLQuery: query,
	  async: false,
      completefunc: function (xData, Status) {
			if (Status == "success") {				
				var collObj = [];
				$(xData.responseXML).SPFilterNode("z:row").each(function() {
					var obj = getAttributes($(this));
					collObj.push(obj);
				});
				collListItem = collObj;
				def.resolve(collListItem);
			}
			else {
				var spErrorcode =  $(xData.responseXML).SPFilterNode("ErrorCode").text(); // 0x00000000 
				var spErrortext =  $(xData.responseXML).SPFilterNode("ErrorText").text() + $(xData.responseXML).SPFilterNode("errorstring").text();
				def.reject('Error in SPServices call\n SharePoint errorcode=' + spErrorcode + '\n ErrorText=' + spErrortext + '\n SPServices status=' + Status);
			}
		}
    });
	return def;
}

function getListItemByTitle(title, list){
	var def = new $.Deferred();			
    var webURL = $().SPServices.SPGetCurrentSite();    	
	var collListItem = [];
	return def;
}

function getAllListItem(list){
	function retrieve(listName, collListItem, startId){
		var id = startId;
		var totalItems = 0;
		var rowLimit = 4999;
		var query = "<Query><Where><Geq><FieldRef Name='ID' /><Value Type='Counter'>"+id+"</Value></Geq></Where></Query>";
		var call = false;

		$().SPServices({
		webURL: $().SPServices.SPGetCurrentSite(),
		operation: 'GetListItems',
		listName: listName,	  
		CAMLQuery: query,
		CAMLRowLimit: rowLimit,
		async: false,
		completefunc: function (xData, Status) {
				if (Status == "success") {
					var collObj = [];
					$(xData.responseXML).SPFilterNode("z:row").each(function() {
						var obj = getAttributes($(this)); // Creates a new object
						collObj.push(obj);
					});
					if(collObj.length > 0){
						call = true;
						collListItem =  $.merge(collListItem, collObj);
						totalItems = collListItem.length;										
						return retrieve(listName, collListItem, (id + rowLimit));
					}				
				}
				else {
					var spErrorcode =  $(xData.responseXML).SPFilterNode("ErrorCode").text(); // 0x00000000 
					var spErrortext =  $(xData.responseXML).SPFilterNode("ErrorText").text() + $(xData.responseXML).SPFilterNode("errorstring").text();
					console.log('Error in SPServices call\n SharePoint errorcode=' + spErrorcode + '\n ErrorText=' + spErrortext + '\n SPServices status=' + Status);
				}
			}
		});
		return collListItem;
	}

    var def = new $.Deferred();
	var objs = retrieve(list, [], 1);
	if(objs != null){
		def.resolve(objs);
	}
	else{
		var spErrorcode =  $(xData.responseXML).SPFilterNode("ErrorCode").text(); // 0x00000000 
		var spErrortext =  $(xData.responseXML).SPFilterNode("ErrorText").text() + $(xData.responseXML).SPFilterNode("errorstring").text();
		def.reject('Error in SPServices call\n SharePoint errorcode=' + spErrorcode + '\n ErrorText=' + spErrortext + '\n SPServices status=' + Status);
	}
    return def;
}

function getListItemByCaml(query, rowLimit, list){
	var def = new $.Deferred();    
	var listItem = null;
	var collListItem = [];		
	
	$().SPServices({
	  webURL: $().SPServices.SPGetCurrentSite(),
      operation: 'GetListItems',
      listName: list,
	  CAMLQuery: query,
	  CAMLRowLimit: rowLimit,
	  CAMLViewFields: '',
	  async: false,
      completefunc: function (xData, Status) {
			if (Status == "success") {
				var collObj = [];
				$(xData.responseXML).SPFilterNode("z:row").each(function() {
					var obj = getAttributes($(this));
					collObj.push(obj);
				});
				collListItem = collObj;
				def.resolve(collListItem);
			}
			else {
				var spErrorcode =  $(xData.responseXML).find("ErrorCode").text(); // 0x00000000 
				var spErrortext =  $(xData.responseXML).find("ErrorText").text() + $(xData.responseXML).find("errorstring").text();
				def.reject('Error in SPServices call\n SharePoint errorcode=' + spErrorcode + '\n ErrorText=' + spErrortext + '\n SPServices status=' + Status);
			}
		}
    });
	return def;
}

// Function to return all of the choices from the Choice column passed in
function getUniqueColumnValuesFromChoice(list, columnName){	
	var def = new $.Deferred();			    
	var choices = [];
    $().SPServices({
	  webURL: $().SPServices.SPGetCurrentSite(),
      operation: 'GetList',
      listName: list,
	  async: false,
      completefunc: function (xData, Status) {
			if (Status == "success") {
				var tempChoices = [];
				$(xData.responseXML).SPFilterNode("Field[DisplayName='" + columnName + "'] CHOICE").each(function () {
					tempChoices.push($(this).text());
				});					
				choices = tempChoices;
				def.resolve(choices);
			}
			else {
				var spErrorcode =  $(xData.responseXML).SPFilterNode("ErrorCode").text(); // 0x00000000 
				var spErrortext =  $(xData.responseXML).SPFilterNode("ErrorText").text() + $(xData.responseXML).SPFilterNode("errorstring").text();
				def.reject('Error in SPServices call\n SharePoint errorcode=' + spErrorcode + '\n ErrorText=' + spErrortext + '\n SPServices status=' + Status);
			}
		}
    });
	return def;
}

function getUserByLogin(login){
	var def = $.Deferred();
	var user = null;
	$().SPServices({
		operation: "GetUserInfo",
		async: false,
		userLoginName: login,
		completefunc: function (xData, Status) {
			if (Status == "success") {						
				$(xData.responseXML).SPFilterNode("User").each(function() {
					user = getAttributes($(this));
					def.resolve(user);
				});								
			}
			else {
				var spErrorcode =  $(xData.responseXML).SPFilterNode("ErrorCode").text(); // 0x00000000
				var spErrortext =  $(xData.responseXML).SPFilterNode("ErrorText").text() + $(xData.responseXML).SPFilterNode("errorstring").text();
				def.reject('Error in SPServices call\n SharePoint errorcode=' + spErrorcode + '\n ErrorText=' + spErrortext + '\n SPServices status=' + Status);
			}			
		}
	});
	return def;
}

function IsCurrentUserMemberOfGroup(groupName) {
	var d = new $.Deferred();
	var currentContext = new SP.ClientContext(_spPageContextInfo.webServerRelativeUrl);//new SP.ClientContext.get_current();
    var currentWeb = currentContext.get_web();
    var currentUser = currentContext.get_web().get_currentUser();
	
    currentContext.load(currentUser);

    var allGroups = currentWeb.get_siteGroups();
    currentContext.load(allGroups);
    currentContext.load(allGroups, 'Include(Users)');
    currentContext.executeQueryAsync(OnSuccess,OnFailure);

    function OnSuccess(sender, args) {
        var userInGroup = false;
        var groupEnumerator = allGroups.getEnumerator();
        while (groupEnumerator.moveNext() && !userInGroup) {
            var oGroup = groupEnumerator.get_current();
            if (oGroup.get_title() == groupName) {
                var collUser = oGroup.get_users();
                var userEnumerator = collUser.getEnumerator();
                while (userEnumerator.moveNext() && !userInGroup) {
                    var oUser = userEnumerator.get_current();
                    if (oUser.get_id() == currentUser.get_id()) {
                        userInGroup = true;
                        break;
                    }
                }
            }
        }
		d.resolve(userInGroup);        
    }

    function OnFailure(sender, args) {
        d.reject('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
    }
	
	return d;
}

function CheckCurrentUserMembership(groupName) {
	var d = new $.Deferred();
	var clientContext = new SP.ClientContext(_spPageContextInfo.webServerRelativeUrl);//SP.ClientContext.get_current();
	this.currentUser = clientContext.get_web().get_currentUser();
	clientContext.load(this.currentUser);
	this.userGroups = this.currentUser.get_groups();
	clientContext.load(this.userGroups);
	currentContext.executeQueryAsync(OnQuerySucceeded, OnQueryFailed);
	
	function OnQuerySucceeded() {
		var isMember = false;
		var groupsEnumerator = this.userGroups.getEnumerator();
		while (groupsEnumerator.moveNext()) {
			var group= groupsEnumerator.get_current();
			if(group.get_title() == groupName) {
				isMember = true;				
				break;
			}
		}
		d.resolve(isMember);	
	}
	
	function OnQueryFailed(sender, args) {
		d.reject('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
	}
	
	return d;
}

function isMemberofSharepointGroup(whatgroup){
	var d = new $.Deferred();
	$().SPServices({
		operation: "GetGroupCollectionFromUser",
		userLoginName: $().SPServices.SPGetCurrentUser(),
		async: false,
		completefunc: function (xData, Status) {
			if(Status === 'success'){
				if ($(xData.responseXML).find("Group[Name='" + whatgroup + "']").length == 1) {
					d.resolve(true);
				}
				else{					
					d.resolve(false);
				}
			}
			else{
				var $xml = $(xData.responseXML); 
				var spErrorcode = $xml.find("ErrorCode").text(); // 0x00000000 
				var spErrortext = $xml.find("ErrorText").text() + $xml.find("errorstring").text();	
				d.reject('Error in SPServices call\n SharePoint errorcode=' + spErrorcode + '\n ErrorText=' + spErrortext + '\n SPServices status=' + Status);
			}
		}
	});
	return d;
}

function retrieveAllGroups() {
	var d = new $.Deferred();
    var clientContext = new SP.ClientContext.get_current();
    this.collGroup = clientContext.get_web().get_siteGroups();
    clientContext.load(collGroup);    

    clientContext.executeQueryAsync(Function.createDelegate(this, function(args, sender){
		var groups = [];
		var groupEnumerator = collGroup.getEnumerator();
		while (groupEnumerator.moveNext()) {
			var oGroup = groupEnumerator.get_current();
			var collUser = oGroup.get_users();
			groups.push(oGroup.get_title());
		}
		d.resolve(groups);
	}), Function.createDelegate(this, function(args, sender){		
		d.reject('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
	}));
	
	return d;
}

function getListItemsByRest(odataQuery){
	var d = new $.Deferred();	
	//"http://sp05.brasilcenter.com.br/sites/InfoCentro/_vti_bin/ListData.svc/ContratoMailing?$expand=ID_CARGA_LOOKUP&$filter=(STATUS_CONTRATOValue%20eq%20%27A%20TRATAR%27%20and%20ASSOCIADO_AId%20eq%20null)&orderby%20ID_CARGA_LOOKUP/DT_INI_TRAT,ID_CARGA_LOOKUP/CD_MAILING"
	$.getJSON(odataQuery)
	.done(function(data){
		var oListItem = typeof(data.d.results) === 'undefined' ? data.d : data.d.results;
		d.resolve(oListItem);
	})
	.fail(function() {
		console.log("error");
	});
	return d.promise();
}

function getUrlVars(url)
{
    var hash;
	var parameters = {}
	hash = url.split('?');
	parameters['$url'] = hash[0];
	parameters['queryString'] = hash[1];

    var hashes = url.slice(url.indexOf('?') + 1).split('&');
    for(var i = 0; i < hashes.length; i++)
    {
        hash = hashes[i].split('=');        
		parameters[hash[0]] = hash[1];
    }
    return parameters;
}

/**API JSOM for manager List in Sharepoint */
function retrieveFieldsOfList(listTitle){

   var context = new SP.ClientContext.get_current();
   var web = context.get_web();
   var list = web.get_lists().getByTitle(listTitle);   
   var listFields = list.get_fields();
   this.fieldsConfig = [];
   context.load(listFields);
   context.executeQueryAsync(printFieldNames,onError);

   function printFieldNames() {
      var e = listFields.getEnumerator();
      while (e.moveNext()) {
         var field = e.get_current();
		if(field.get_canBeDeleted()){
			fieldsConfig.push({
				title: field.get_title(),
				internalName: field.get_internalName(),
				type: field.get_typeAsString(), 
				IsNodeDevice: false,
				DependentOfField: '',
				DependentOfValue: ''
			});			
		}
      }

	  console.log(fieldsConfig);
   }

   function onError(sender,args)
   {
      console.log(args.get_message());
   }
}

function retrieveListItems(caml, listTitle, strInclude) {
	var d = new $.Deferred();

	SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function(){
		var clientContext = new SP.ClientContext.get_current();
		var oList = clientContext.get_web().get_lists().getByTitle(listTitle);
		var listFields = oList.get_fields();
		var oItems = [];
		var camlQuery = new SP.CamlQuery();
		//<View> <ViewFields> <FieldRef Name="Title"/> <FieldRef Name="STATUS_CONTRATO"/> <FieldRef Name="MailingInfoID_CARGA"/> <FieldRef Name="MailingInfoCD_MAILING"/> <FieldRef Name="MailingInfoDT_INI_TRAT"/> </ViewFields> <Joins> <Join Type="INNER" ListAlias="MailingInfo"> <Eq> <FieldRef Name="ID_CARGA_LOOKUP" RefType="ID"/> <FieldRef Name="ID" List="MailingInfo"/> </Eq> </Join> </Joins> <ProjectedFields> <Field ShowField="CD_MAILING" Type="Lookup" Name="MailingInfoCD_MAILING" List="MailingInfo"/> <Field ShowField="ID_CARGA" Type="Lookup" Name="MailingInfoID_CARGA" List="MailingInfo"/> <Field ShowField="DT_INI_TRAT" Type="Lookup" Name="MailingInfoDT_INI_TRAT" List="MailingInfo"/> </ProjectedFields> <Query> <Where> <Eq> <FieldRef Name="STATUS_CONTRATO"/> <Value Type="Choice">A TRATAR</Value> </Eq> </Where> <OrderBy> <FieldRef Name="MailingInfoDT_INI_TRAT"/> <FieldRef Name="MailingInfoCD_MAILING"/> </OrderBy> </Query></View>
		camlQuery.set_viewXml(caml);
		this.collListItem = oList.getItems(camlQuery);

		if(strInclude !== '' || typeof(strInclude) !== 'undefined')
			//'Include(Id, Title, STATUS_CONTRATO, MailingInfoCD_MAILING, MailingInfoDT_INI_TRAT)'
			clientContext.load(collListItem, strInclude);
		else
			clientContext.load(collListItem);

		clientContext.load(listFields);
		clientContext.executeQueryAsync(
			Function.createDelegate(this, onQuerySucceeded), 
			Function.createDelegate(this, onQueryFailed));

		function onQuerySucceeded(sender, args) {		
			var listItemEnumerator = collListItem.getEnumerator();
				
			while (listItemEnumerator.moveNext()) {
				var oListItem = listItemEnumerator.get_current();
				var fieldEnumerator = listFields.getEnumerator();
				var oItem = {};
				while(fieldEnumerator.moveNext()){
					var field = fieldEnumerator.get_current();
					if(field.get_canBeDeleted()){
						defineProp(oItem, field.get_internalName, oListItem[field.get_internalName]);
					}
				}
				oItems.push(oItem);
			}
			d.resolve(oItems);
		}

		function onQueryFailed(sender, args) {		
			d.reject('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());			
		}
	});    
	return d.promise();
}

function retrieveAllListProperties() {
	var d = new $.Deferred();
	SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function(){
		var clientContext = new SP.ClientContext.get_current();
		var oWebsite = clientContext.get_web();
		this.collList = oWebsite.get_lists();
		//clientContext.load(collList, 'Include(Fields)');
		this.listInfoArray = clientContext.loadQuery(collList, 
        'Include(Title,Created,LastItemModifiedDate,ItemCount,BaseTemplate,BaseType,Fields.Include(Title,InternalName,CanBeDeleted))');
		clientContext.executeQueryAsync(
			Function.createDelegate(this, onQuerySucceeded), 
			Function.createDelegate(this, onQueryFailed)
		);

		function onQuerySucceeded() {
			var listInfo = [];

			for (var i = 0; i < this.listInfoArray.length; i++) {            
				var oList = this.listInfoArray[i];
				var collField = oList.get_fields();					
				var fieldEnumerator = collField.getEnumerator();
				var templateId = oList.get_baseTemplate();
				var baseId = oList.get_baseType();
				var baseTemplate = '';
				var baseType = '';

				switch (baseId) {
					case 0:
						baseType = 'Custom List';
						break;
					
					case 1:
						baseType = 'Document Library';
						break;
					
					case 2:
						baseType = 'Not used';
						break;

					case 3:
						baseType = 'Obsolete. Use 0 for discussion boards.';
						break;

					case 4:
						baseType = 'Surveys';
						break;
					
					case 5:
						baseType = 'Issues List';
						break;
				}
				switch (templateId) {
					case 100:
						baseTemplate = 'Custom list';
						break;

					case 101:
						baseTemplate = 'Document library';
						break;
					
					case 102:
						baseTemplate = 'Survey';
						break;

					case 103:
						baseTemplate = 'Links';
						break;

					case 104:
					baseTemplate = 'Announcements';
					break;

					case 105:
					baseTemplate = 'Contacts';
					break;

					case 106:
					baseTemplate = 'Calendar';
					break;

					case 107:
					baseTemplate = 'Tasks';
					break;

					case 108:
					baseTemplate = 'Discussion board';
					break;

					case 109:
					baseTemplate = 'Picture library';
					break;

					case 110:
					baseTemplate = 'Data sources for a site';
					break;

					case 111:
					baseTemplate = 'Site template gallery';
					break;

					case 112:
					baseTemplate = 'User Information';
					break;

					case 113:
					baseTemplate = 'Web Part gallery';
					break;

					case 114:
					baseTemplate = 'List Template gallery';
					break;

					case 115:
					baseTemplate = 'XML Form library';
					break;

					case 116:
					baseTemplate = 'Master Page gallery';
					break;

					case 117:
					baseTemplate = 'No Code Workflows';
					break;

					case 118:
					baseTemplate = 'Custom Workflow Process';
					break;

					case 119:
					baseTemplate = 'Wiki Page Library';
					break;

					case 120:
					baseTemplate = 'Custom grid for a list';
					break;

					case 121:
					baseTemplate = 'Solutions';
					break;

					case 122:
					baseTemplate = 'No Code Public Workflow';
					break;

					case 123:
					baseTemplate = 'Themes';
					break;

					case 130:
					baseTemplate = 'DataConnectionLibrary';
					break;

					case 140:
					baseTemplate = 'Workflow History';
					break;

					case 150:
					baseTemplate = 'Project Tasks';
					break;

					case 200:
					baseTemplate = 'Meeting Series (Meeting)';
					break;

					case 201:
					baseTemplate = 'Agenda (Meeting)';
					break;

					case 202:
					baseTemplate = 'Attendees (Meeting)';
					break;

					case 204:
					baseTemplate = 'Decisions (Meeting)';
					break;

					case 207:
					baseTemplate = 'Objectives (Meeting)';
					break;

					case 210:
					baseTemplate = 'Text Box (Meeting)';
					break;

					case 211:
					baseTemplate = 'Things To Bring (Meeting)';
					break;

					case 212:
					baseTemplate = 'Workspace Pages (Meeting)';
					break;

					case 301:
					baseTemplate = 'Posts (Blog)';
					break;

					case 302:
					baseTemplate = 'Comments (Blog)';
					break;

					case 303:
					baseTemplate = 'Categories (Blog)';
					break;

					case 402:
					baseTemplate = 'Facility';
					break;

					case 403:
					baseTemplate = 'Whereabouts';
					break;

					case 404:
					baseTemplate = 'Call Track';
					break;

					case 405:
					baseTemplate = 'Circulation';
					break;

					case 420:
					baseTemplate = 'Timecard';
					break;

					case 421:
					baseTemplate = 'Holidays';
					break;

					case 499:
					baseTemplate = 'ME (Input Method Editor) Dictionary';
					break;

					case 600:
					baseTemplate = 'External';
					break;

					case 1100:
					baseTemplate = 'Issue tracking';
					break;

					case 1200:
					baseTemplate = 'Administrator Tasks';
					break;

					case 1220:
					baseTemplate = 'Health Rules';
					break;

					case 1221:
					baseTemplate = 'Health Reports';
					break;
				}
				var info = {
					Title: oList.get_title(),
					Created: oList.get_created().toLocaleString(),
					Modified: oList.get_lastItemModifiedDate().toLocaleString(),
					ItemCount: oList.get_itemCount(),
					baseTemplate: baseTemplate,
					baseType: baseType
				}
				var fieldsCanBeDeleted = [];
				while (fieldEnumerator.moveNext()) {
					field = fieldEnumerator.get_current();
					if(field.get_canBeDeleted()){						
						fieldsCanBeDeleted.push({
								internalName: field.get_internalName(), 
								DisplayName: field.get_title()
							}); 
					}
				}
				defineProp(info, 'fieldsCanBeDeleted', fieldsCanBeDeleted);
				listInfo.push(info);
			}
			d.resolve(listInfo);
		}

		function onQueryFailed(sender, args) {
			d.reject('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
		}
	});
	return d.promise();
}

function getAllListInfo(){
	var deferred = new $.Deferred();
	$().SPServices({
		webURL: $().SPServices.SPGetCurrentSite(),
		operation: "GetListCollection",
		completefunc: function(xData, Status) {
			if(Status === 'success'){
				var collObj = [];	
				
				var obj = $(xData.responseXML).SPFilterNode("List").SPXmlToJson({
					mapping: {},
					includeAllAttrs: false,
					removeOws: true,
					sparse: false // Added in 2014.01
				});
				deferred.resolve(obj);
			}
			else{
				var spErrorcode =  $(xData.responseXML).find("ErrorCode").text(); // 0x00000000 
				var spErrortext =  $(xData.responseXML).find("ErrorText").text() + $(xData.responseXML).find("errorstring").text();
				deferred.reject('Error in SPServices call\n SharePoint errorcode=' + spErrorcode + '\n ErrorText=' + spErrortext + '\n SPServices status=' + Status);				
			}
		}
	});

	return deferred.promise();
}

function getListInfo(listName){
	var d = new $.Deferred();
	SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function(){
		var clientContext = new SP.ClientContext.get_current();
		var oWebsite = clientContext.get_web();
		var oList = oWebsite.get_lists().getByTitle(listName);
		var collFields = oList.get_fields();
		clientContext.load(oList);
		clientContext.load(collFields);
		clientContext.executeQueryAsync(
			Function.createDelegate(this, onQuerySucceeded), 
			Function.createDelegate(this, onQueryFailed)
		);

		function onQuerySucceeded() {
			var info = {
				Title: oList.get_title(),
				Created: oList.get_created().toLocaleString(),
				Modified: oList.get_lastItemModifiedDate().toLocaleString(),
				ItemCount: oList.get_itemCount(),
			}

			var fieldEnumerator = collFields.getEnumerator();
			var fieldsCanBeDeleted = [];
			while(fieldEnumerator.moveNext()){
                field = fieldEnumerator.get_current();
                if(field.get_canBeDeleted()){
					fieldsCanBeDeleted.push({internalName: field.get_internalName(), DisplayName: field.get_title()}); 
                }
            }
			defineProp(info, 'fieldsCanBeDeleted', fieldsCanBeDeleted);
			d.resolve(info);
		}
		function onQueryFailed(sender, args) {
			d.reject('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
		}
	});
	return d.promise();
}

function deleteListItemDisplayCount(listName, items) {
	var deferred = new $.Deferred();
	this.allItems = items;
	SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function(){
		this.clientContext = new SP.ClientContext.get_current();
		this.oList = clientContext.get_web().get_lists().getByTitle(listName);

		var o = {deferred: deferred, allItems:allItems};
		clientContext.load(oList);
		clientContext.executeQueryAsync(
			Function.createDelegate(o, deleteItem), 
			Function.createDelegate(o, onQueryFailed)
		);

		function deleteItem() {
			this.itemId = -1;
			this.startCount = -1;
			this.oListItem = null;			
			this.index = 0;

			this.allItems.forEach(function(item){				
				startCount = oList.get_itemCount();
				oListItem = oList.getItemById(item.ows_ID);
				oListItem.deleteObject();			
				
				if(index % 25 == 0){
					oList.update();
					clientContext.load(oList);
					clientContext.executeQueryAsync(
						Function.createDelegate(o, displayCount),
						Function.createDelegate(o, onQueryFailed)
					);
				}
				index++;
			});
			clientContext.executeQueryAsync(
				Function.createDelegate(o, displayCount),
				Function.createDelegate(o, onQueryFailed)
			);			
		}

		function displayCount() {
			var endCount = oList.get_itemCount();
			var listItemInfo = 'Start Count: ' +  startCount + ' End Count: ' + endCount;
			console.log(listItemInfo);
			this.deferred.resolve(endCount);
		}

		function onQueryFailed(sender, args) {
			this.deferred.reject('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
		}
	});

	return deferred.promise();
}