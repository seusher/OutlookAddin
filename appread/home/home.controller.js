(function() {
    'use strict';

    angular.module('officeAddin')
        .controller('homeController', ['dataService', '$http', '$q', homeController]);

    /**
     * Controller constructor
     */
    function homeController(dataService, $http, $q) {
        var vm = this;  // jshint ignore:line

        vm.emails = {};
        vm.image = "";

        Office.initialize = function() {
            //console.log('>>> Office.initialize()');
            getDataFromService();
        };
        getDataFromService();

        function getGroupData(groupId)
        {
            var deferred = $q.defer();

            $http({
                  method: 'GET',
                  url: 'https://graph.microsoft.com/v1.0/groups/' + groupId,
                }).then(function(response) {

                  var groupDetails = response;
                  var result = {
                      displayName:groupDetails.data.displayName,
                      description:groupDetails.data.description,
                      image:null};

                  console.log('Got group: ' + groupDetails.data.displayName);
                  if (response.data.visibility != null) {

                    // Add the image to the response
                    getGroupImage(groupId).then(function(response) {

                      // Since an image was found, add it to the result
                      result.image = response;

                      deferred.resolve(result);
                    }, function(error) {
                      console.log('Failed to get the image.');
                    });
                  }
                  else {
                    // If the group has a null visibility, assume it doesn't
                    // have an photo and skip the photo lookup.
                    deferred.resolve(result);
                  }

                }, function(error) {
                  console.log('Failed to Get Group ' + groupId);
                });

            return deferred.promise;
        }

        // Get the group image.
        // It will be returned as a base64-encded string
        function getGroupImage(groupId)
        {
            var deferred = $q.defer();
            $http({
                method: 'GET',
                url: 'https://graph.microsoft.com/v1.0/groups/' + groupId + '/photo/$value',
                responseType: 'blob',
                headers: {
                  'Content-Type': 'application/octet-binary'
                }
              })
            .then(function(response) {

              // Use FileReader to convert the image from a blob to a
              // base64-encoded string.
              var reader = new FileReader();
              reader.onloadend = function() {
                deferred.resolve(reader.result);
              }
              reader.readAsDataURL(response.data);

            }, function(error) {
              console.log('HTTP request to Group Photo API failed.');
            });

            return deferred.promise;
        }

        function getDataFromService() {

          // Querying for all groups requires the 'Read all groups' permission to be enabled for your app.
          // This permission has to be approved by the tenant administrator.
          $http({
                method: 'GET',
                url: 'https://graph.microsoft.com/v1.0/groups?$top=100',
              })
            .then(function(response) {

                  // At this point we have the set of groups that I belong to
                  var arr = [];
                  for (var i = 0, len = response.data.value.length; i < len; i++) {
                      // Look up the actual group details for each group
                      arr.push(getGroupData(response.data.value[i].id));
                  }

                  $q.all(arr).then(function (ret) {
                      // ret[0] contains the response of the first call
                      // ret[1] contains the second response
                      // etc.

                      var results = []
                      for (var i = 0, len = ret.length; i < len; i++) {

                        if (ret[i].image != null){
                          results.push(ret[i]);
                        }
                      }

                      vm.groups = results;
                  });

            }, function(error) {
              console.log('HTTP request to Graph API failed.');
            });
        }
    }

})();
