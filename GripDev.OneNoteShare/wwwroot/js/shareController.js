// Copyright (c) Microsoft 2015. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information. 

'use strict';

var AppUrl = 'YOUR-APPS-URL-HERE'
var ClientId = 'YOUR-CLIENT-ID-HERE'

// module
var App = angular.module('app', []);

// controller
App.controller('shareController', function ($scope, $q, $http, $window) {
    var vm = this;

    vm.title = "test";
    vm.url = $window.location.toString();
    vm.linkToOneNotePage = null;


    //When the page loads look to see if an access token is included in the URL
    //Warning - Simple positional parsing of the access_token parameter here as an example, replace for production with robust parsing. 
    if ($window.location.hash.indexOf('=') === -1) {
        vm.authToken = null;
    }
    else {
        vm.authToken = $window.location.hash.substring($window.location.hash.indexOf('=') + 1, $window.location.hash.length);
        vm.authToken = vm.authToken.substring(0, vm.authToken.indexOf('&'));
    }

    vm.export = function () {

        //Navigate to the authorization page, without popup, by changing the window location
        var clientId = ClientId;
        var redirectUrl = encodeURIComponent(AppUrl)

        var url = 'https://login.live.com/oauth20_authorize.srf?client_id='
            + clientId
            + '&display=page&locale=en&redirect_uri='
            + redirectUrl + '&response_type=token&scope=wl.signin%20office.onenote_create&response_type=token';

        $window.location.href = url;
    };

    vm.postPage = function () {
        var that = this;
        return $http.post(
          'https://www.onenote.com/api/v1.0/pages',
          '<html><head><title>Hello</title></head>' +
            new Date().toLocaleString() +
          '</html>',
          {
              headers: {
                  'Content-Type': 'text/html',
                  'Authorization': 'Bearer ' + vm.authToken
              }
          }
        )
        .then(function (response) {
            that.linkToOneNotePage = response.data.links.oneNoteWebUrl.href;
        });
    }
});