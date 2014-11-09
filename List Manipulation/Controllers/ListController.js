(function () {
    //"use strict";
    var app = angular.module("app1", [
        // angular stuff
        "ng",
        "ui.router",
        "sharePointList",
        "ngGrid",
        'ngSanitize'
    ]);
    app.config(function (SharePointListProvider, $stateProvider, $urlRouterProvider) {
        // For any unmatched url, redirect to /Employee
        $urlRouterProvider.otherwise("/Employee");
        $stateProvider
         .state('Employee', {
             url: '/Employee',
             templateUrl: 'ListView.html',
             controller: "ListController"
         });
    });
    app.controller("SharePointListCtrl", ["$scope", "$state", "SharePointList", "$templateCache", function ($scope, $state, SharePointList, $templateCache) {
        $scope.gridOptions = {
            data: 'Employee',
            enableCellSelection: true,
            enableRowSelection: false,
            enableCellEdit: true,
            columnDefs: [
                { field: 'Title', displayName: 'Title', enableCellEdit: false, width: "160px", resizeable: true },
                { field: 'RoleId', displayName: 'Role', width: "160px" }
            ]
        };
        var queryString = "<View><Query><OrderBy><FieldRef Name='Title' Ascending='False'/></OrderBy></Query></View>";
        SharePointList.getListItems("Employee", queryString, ["Title", "RoleId"]).then(function (Employee) {
            $scope.Employees = Employee;
        });
    }]);
})();
window.onload = function () {
    var $rootelement = angular.element(window.document);
    var modules = [
        "ng",
        "app1",
        function ($provide) { $provide.value("$rootElement", $rootelement); }
    ];
    
    var $injector = angular.injector(modules);
    var $compile = $injector.get("$compile");
    var compositeLinkFn = $compile($rootelement);
    var $rootScope = $injector.get("$rootScope");
    compositeLinkFn($rootScope);
    $rootScope.$apply();
};