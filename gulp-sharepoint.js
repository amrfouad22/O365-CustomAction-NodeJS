var request = require('request');
var sharepoint = require('sharepoint-apponly-node');
var NodeCache = require("node-cache");
var cache = new NodeCache();

module.exports = {
    getToken: function (siteUrl, clientId, clientSecret, callback) {
        cache.get('token', function (err, value) {
            if (!err && value != undefined) {
                return callback(value);
            }
            sharepoint.getSharePointAppOnlyAccessToken(siteUrl, clientId, clientSecret, function (response) {
                cache.set('token', response.access_token, response.expires_in, function (err, success) {
                    callback(response.access_token);
                })
            });
        })
    },
    getCustomAction: function (web,guid,token, callback) {
        if(!guid){
            spGet(web+'/_api/web/usercustomactions',token,callback);
        }
        else{
            spGet(web+'/_api/web/usercustomactions(guid\''+guid+'\')',token,callback);            
        }
    },
    deleteCustomAction:function(web,guid,token,callback){
        if(!guid){
            throw 'please provide custom action guid'
        }
        spDelete(web+'/_api/web/usercustomactions(guid\''+guid+'\')',token,callback);
    },
    createCustomAction:function(web,body,token,callback){
        if(!body){
            throw 'please provide custom action properties'
        }
        spPost(web+'/_api/web/usercustomactions',body,token,callback);
    }
    
}
function spPost(url,body,token,callback){
    execRequest(url,body,'POST',token,callback);
}
function spDelete(url,token,callback){
    execRequest(url,null,'DELETE',token,callback);
}
function spGet(url,token,callback){
    execRequest(url,null,'GET',token,callback);
}

function execRequest(url,body,method,token,callback){
     request(
            {
                url: url,
                method: method,
                headers: {
                    'Authorization': 'Bearer ' + token,
                    'Content-Type': 'application/json',
                    'Accept': 'application/json;odata=verbose',
                },
                json: true,
                body:body
            },function(err,response,body){
                if(err){
                    throw 'error executing request '+method+'to '+url;
                }
                console.log(body);
                callback(body);
            });
}