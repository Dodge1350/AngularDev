// SharePointListService.js
(function () {
    "use strict";
    var module = angular.module('sharePointList', []);
    module.provider('SharePointList', function () {
        var clientCtx;
        var web;
        var configuration = {};
        this.$get = ['$q', "$log", function ($q, $log) {
            var contextLoaded = $q.defer();
            SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
                clientCtx = SP.ClientContext.get_current();
                web = clientCtx.get_web();
                contextLoaded.resolve();
            });
            function createServiceForConfiguration(config) {
                var service = {};
                service.self = service;
                service.clientCtx = clientCtx;
                service.web = web;
                service.getListItems = function (listName, queryString, fields) {
                    var deferred = $q.defer();
                    contextLoaded.promise.then(function () {
                        var list = web.get_lists().getByTitle(listName);
                        var query = new SP.CamlQuery();
                        query.set_viewXml(queryString);
                        var listItems = list.getItems(query);
                        var fieldList = fields.join(",");
                        clientCtx.load(listItems, "Include(" + fieldList + ")");
                        clientCtx.executeQueryAsync(function () {
                            var resultItems = [];
                            var listItemEnumerator = listItems.getEnumerator();
                            while (listItemEnumerator.moveNext()) {
                                var listItem = listItemEnumerator.get_current();
                                var resultItem = {};
                                for (i = 0; i < fields.length; i++) {
                                    resultItem[fields[i]] = listItem.get_item(fields[i]);
                                }
                                resultItems.push(resultItem);
                            }
                            deferred.resolve(resultItems);
                        }, function (sender, args) {
                            var messageFormat = "Loading of list {0} with  failed with error. {1} \n{2}";
                            var message = messageFormat.format(listName, args.get_message(), args.get_stackTrace());
                            $log.error(message);
                        })
                    });
                    return deferred.promise;
                };
                service.createListItem = function (listName, item, fields) {
                    var deferred = $q.defer();
                    contextLoaded.promise.then(function () {
                        var oList = web.get_lists().getByTitle(listName);
                        var itemCreateInfo = new SP.ListItemCreationInformation();
                        this.oListItem = oList.addItem(itemCreateInfo);
                        for (var property in item) {
                            var includeField;
                            if (!fields) {
                                includeField = true;
                            }
                            else {
                                includeField = false;
                                for (var i = 0; i < fields.length; i++) {
                                    if (fields[i] === property) {
                                        includeField = true;
                                        break;
                                    }
                                }
                            }
                            if (item.hasOwnProperty(property) && includeField) {
                                oListItem.set_item(property, item[property]);
                            }
                        }
                        oListItem.update();
                        clientCtx.load(oListItem);
                        clientCtx.executeQueryAsync(function () {
                            deferred.resolve(oListItem.get_fieldValues());
                        }, function (sender, args) {
                            var messageFormat = "Loading of list {0} with  failed with error. {1} \n{2}";
                            var message = messageFormat.format(listName, args.get_message(), args.get_stackTrace());
                            $log.error(message);
                        })
                    });
                    return deferred.promise;
                };
                service.deleteListItem = function (listName, itemId) {
                    var deferred = $q.defer();
                    contextLoaded.promise.then(function () {
                        var oList = web.get_lists().getByTitle(listName);
                        this.oListItem = oList.getItemById(itemId);
                        oListItem.deleteObject();
                        clientCtx.executeQueryAsync(function () {
                            deferred.resolve();
                        }, function (sender, args) {
                            var messageFormat = "deleting item {0} from  list {1} with  failed with error. {2} \n{3}";
                            var message = messageFormat.format(itemId, listName, args.get_message(), args.get_stackTrace());
                            $log.error(message);
                        })
                    });
                    return deferred.promise;
                };
                service.saveListItem = function (listName, item, fields) {
                    var deferred = $q.defer();
                    contextLoaded.promise.then(function () {
                        var oList = web.get_lists().getByTitle(listName);
                        var oListItem = oList.getItemById(item.ID);
                        for (var property in item) {
                            var includeField;
                            if (!fields) {
                                includeField = true;
                            }
                            else {
                                includeField = false;
                                for (var i = 0; i < fields.length; i++) {
                                    if (fields[i] === property) {
                                        includeField = true;
                                        break;
                                    }
                                }
                            }
                            if (item.hasOwnProperty(property) && includeField && property != "ID") {
                                oListItem.set_item(property, item[property]);
                            }
                        }
                        oListItem.update();
                        clientCtx.load(oListItem);
                        clientCtx.executeQueryAsync(function () {
                            deferred.resolve(oListItem.get_fieldValues());
                        }, function (sender, args) {
                            var messageFormat = "update of list {0} with  failed with error. {1} \n{2}";
                            var message = messageFormat.format(listName, args.get_message(), args.get_stackTrace());
                            $log.error(message);
                        })
                    });
                    return deferred.promise;
                };
                return service;
            }
            return createServiceForConfiguration(configuration);
        }];
    }
    );
})();