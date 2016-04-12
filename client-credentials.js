//==============================================================================
// Exports the ClientCredentials class that provides the getAccessToken method.
// The mechanism by which the access token is obtained is described by the
// "Service to Service Calls Using Client Credentials" article, available at
// https://msdn.microsoft.com/en-us/library/azure/dn645543.aspx.
//==============================================================================
// Author: Frank Hellwig
// Copyright (c) 2016 Buchanan & Edwards
//==============================================================================

'use strict';

//------------------------------------------------------------------------------
// Dependencies
//------------------------------------------------------------------------------

const HttpsService = require('https-service');
const util = require('util');

//------------------------------------------------------------------------------
// Initialization
//------------------------------------------------------------------------------

const FIVE_MINUTE_BUFFER = 5 * 60 * 1000;
const loginService = new HttpsService('login.microsoftonline.com');

//------------------------------------------------------------------------------
// Public
//------------------------------------------------------------------------------

class ClientCredentials {
    constructor(tenant, clientId, clientSecret) {
        this.base = util.format('/%s/oauth2/', tenant);
        this.clientId = clientId;
        this.clientSecret = clientSecret;
        this.tokens = {};
    }

    getAccessToken(resource, callback) {
        let token = this.tokens[resource];
        if (token) {
            var now = new Date();
            if (now.getTime() < token.exp) {
                return callback(null, token.val);
            }
        }
        this.requestAccessToken(resource, callback);
    }

    requestAccessToken(resource, callback) {
        let body = {
            grant_type: 'client_credentials',
            client_id: this.clientId,
            client_secret: this.clientSecret,
            resource: resource
        };
        let self = this;
        self.httpsPost('token', body, function(err, response) {
            if (err) return callback(err);
            let now = new Date();
            let exp = now.getTime() + (parseInt(response.expires_in) * 1000);
            self.tokens[resource] = {
                val: response.access_token,
                exp: exp - FIVE_MINUTE_BUFFER
            };
            callback(null, response.access_token);
        });
    }

    /**
     * Sends an HTTPS POST request to the specified endpoint.
     * The endpoint is the last part of the URI (e.g., "authorize").
     */
    httpsPost(endpoint, body, callback) {
        let path = this.base + endpoint;
        let headers = {
            'Content-Type': HttpsService.FORM_MEDIA_TYPE
        };
        loginService.request('POST', path, headers, body, callback);
    }
}

//------------------------------------------------------------------------------
// Exports
//------------------------------------------------------------------------------

module.exports = ClientCredentials;
