(function() {
    'use strict';

    angular.module('officeAddin')
        .service('dataService', ['$q', '$http', dataService]);

    /**
     * Custom Angular service.
     */
    function dataService($q, $http) {

        // public signature of the service
        return {
            getData: getData
        };

        /** *********************************************************** */

        function getData() {
            var deferred = $q.defer();
            $http.get('https://graph.microsoft.com/v1.0/me/messages')
                .then(function(results) {
                    deferred.resolve(results);
                });
            return deferred.promise;
        }
    }
})();
