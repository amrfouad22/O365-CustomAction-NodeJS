var sharepoint = require('sharepoint-apponly-node');
var request = require('request');
var O365Constants = require('./O365Constants.js');
var NodeCache = require("node-cache");

var myCache = new NodeCache();
module.exports = {
    acquireToken: function (callback) {
        myCache.get('token', function (error, value) {
            if (error || value == undefined) {
                sharepoint.getSharePointAppOnlyAccessToken(O365Constants.resource, O365Constants.clientId, O365Constants.clientSecret,
                    function (tokenResponse) {
                        myCache.set('token', tokenResponse.access_token,tokenResponse.expires_in);
                        return callback(tokenResponse.access_token);
                    });
                
            }
            return callback(value);
        });
    },
    updateItemTitle: function (listId, itemId, callback) {

        this.acquireToken(function (token) {
            getDigest(token, function (digest) {
                getItemEtag(token, listId, itemId, function (etag) {
                    updateItemTitle(token, listId, itemId, digest.d.GetContextWebInformation.FormDigestValue, etag, callback);
                })
            })
        });
    }
}
function updateItemTitle(token, listId, itemId, digest, etag, callback) {
    request(
        {
            url: O365Constants.resource + '/_api/web/lists(guid\'' + listId + '\')/items(' + itemId + ')',
            method: 'PATCH',
            headers: {
                'Authorization': 'Bearer ' + token,
                'Content-Type': 'application/json',
                "X-RequestDigest": digest,
                'X-HTTP-Method': 'PATCH',
                'Accept': 'application/json;odata=verbose',
                'If-Match': etag
            },
            body: {
                'Title': 'Title updated by NodeJS Custom Action'
            },
            json: true
        })
        .on('error', function (err) {
            callback(err);
        })
        .on('response', function (response) {
            callback(response);
        });

}
function getDigest(token, callback) {
    request(
        {
            url: O365Constants.resource + '/_api/contextinfo',
            method: 'POST',
            headers: {
                'Authorization': 'Bearer ' + token,
                'Accept': 'application/json; odata=verbose'
            },
            json: true
        }, function (error, response, body) {
            if (!error && response.statusCode == 200) {
                callback(body);
            }
        });
}
function getItemEtag(token, listId, itemId, callback) {
    request(
        {
            url: O365Constants.resource + '/_api/web/lists(guid\'' + listId + '\')/items(' + itemId + ')',
            method: 'GET',
            headers: {
                'Authorization': 'Bearer ' + token,
                'Content-Type': 'application/json',
                'Accept': 'application/json;odata=verbose',
            },
            json: true
        }, function (error, response, body) {
            if (!error && response.statusCode == 200) {
                callback(response.headers.etag);
            }
        });
}
