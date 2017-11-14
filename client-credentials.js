//==============================================================================
// Exports the ClientCredentials class that provides the getAccessToken method.
// The mechanism by which the access token is obtained is described by the
// "Service to Service Calls Using Client Credentials" article, available at
// https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-protocols-oauth-service-to-service
//==============================================================================
// Author: Frank Hellwig
// Copyright (c) 2017 Buchanan & Edwards
//==============================================================================

'use strict';

//------------------------------------------------------------------------------
// Dependencies
//------------------------------------------------------------------------------

const request = require('request');

//------------------------------------------------------------------------------
// Initialization
//------------------------------------------------------------------------------

const FIVE_MINUTE_BUFFER = 5 * 60 * 1000;
const MICROSOFT_LOGIN_URL = 'https://login.microsoftonline.com';

//------------------------------------------------------------------------------
// Public
//------------------------------------------------------------------------------

class ClientCredentials {
  constructor(tenant, clientId, clientSecret) {
    this.tenant = tenant;
    this.clientId = clientId;
    this.clientSecret = clientSecret;
    this.tokens = {};
  }

  /**
   * Gets the access token from the login service or returns a cached
   * access token if the token expiration time has not been exceeded. 
   * @param {string} resource - The App ID URI for which access is requested.
   * @returns A promise that is resolved with the access token.
   */
  getAccessToken(resource) {
    let token = this.tokens[resource];
    if (token) {
      var now = new Date();
      if (now.getTime() < token.exp) {
        return Promise.resolve(token.val);
      }
    }
    return this._requestAccessToken(resource);
  }

  /**
   * Requests the access token using the OAuth 2.0 client credentials flow.
   * @param {string} resource - The App ID URI for which access is requested.
   * @returns A promise that is resolved with the access token.
   */
  _requestAccessToken(resource) {
    let params = {
      grant_type: 'client_credentials',
      client_id: this.clientId,
      client_secret: this.clientSecret,
      resource: resource
    };
    return this._httpsPost('token', params).then(body => {
      let now = new Date();
      let exp = now.getTime() + parseInt(body.expires_in) * 1000;
      this.tokens[resource] = {
        val: body.access_token,
        exp: exp - FIVE_MINUTE_BUFFER
      };
      return body.access_token;
    });
  }

  /**
   * Sends an HTTPS POST request to the specified endpoint.
   * The endpoint is the last part of the URI (e.g., "token").
   */
  _httpsPost(endpoint, params) {
    const options = {
      method: 'POST',
      baseUrl: MICROSOFT_LOGIN_URL,
      uri: `/${this.tenant}/oauth2/${endpoint}`,
      form: params,
      json: true,
      encoding: 'utf8'
    };
    return new Promise((resolve, reject) => {
      request(options, (err, response, body) => {
        if (err) return reject(err);
        if (body.error) {
          err = new Error(body.error_description.split(/\r?\n/)[0]);
          err.code = body.error;
          return reject(err);
        }
        resolve(body);
      });
    });
  }
}

//------------------------------------------------------------------------------
// Exports
//------------------------------------------------------------------------------

module.exports = ClientCredentials;
