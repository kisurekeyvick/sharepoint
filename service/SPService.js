var hostweburl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
var spAppWebUrl = decodeURIComponent(getQueryStringParameter("SPAppWebUrl"));

(function () {
    'use strict';

    angular.module('ServiceMe.Service.SharePointService', ['ngResource'])
    .factory('HttpExecutor', HttpExecutor)
    .factory('JSOMExecutor', JSOMExecutor)
    .service('SharePointFileService', SharePointFileService)
    .service('SharePointJSOMService', SharePointJSOMService)
    .factory('SharePointResource', SharePointResource)
    .factory('SharePointAPIService', SharePointAPIService)
    .service('SharePointFormService', SharePointFormService)
    .service('SharePointListService', SharePointListService)
    .service('SharePointPermissionsService', SharePointPermissionsService);

    HttpExecutor.$inject = ['$http'];
    function HttpExecutor($http) {
        return {
            get: function (url, headers) {
                headers || (headers = {});
                if (AppOnlyAccessTokenForSPHost != '')
                    headers['Authorization'] = 'Bearer ' + AppOnlyAccessTokenForSPHost;

                return $http.get(url, { headers: headers });
            },
            post: function (url, data, headers) {
                headers || (headers = {});
                if (AppOnlyAccessTokenForSPHost != '')
                    headers['Authorization'] = 'Bearer ' + AppOnlyAccessTokenForSPHost;
                headers['content-length'] = headers['content-length'] || data.length;

                return $http.post(url, data, { headers: headers });
            },
            postWithoutTransformRequest: function (url, data, headers, userMode) {
                headers || (headers = {});
                if (AppOnlyAccessTokenForSPHost != '') {
                    headers['Authorization'] = 'Bearer ' + AppOnlyAccessTokenForSPHost;
                }
                if (userMode)
                    headers['Authorization'] = 'Bearer ' + UserAccessTokenForSPHost;
                headers['content-length'] = headers['content-length'] || data.length;

                return $http({
                    method: 'POST',
                    url: url,
                    headers: headers,
                    data: data,
                    transformRequest: []
                });
            }
        }
    }

    JSOMExecutor.$inject = [];
    function JSOMExecutor() {
        return function (url, useApp) {
            var hosturl = hostweburl;
            if (url != undefined)
                hosturl = url;
            if (useApp == undefined)
                useApp = true;
            var clientContext;
            var web;
            if (hosturl) {
                if (appMode) {
                    clientContext = new SP.ClientContext(spAppWebUrl);
                    var factory = new SP.ProxyWebRequestExecutorFactory(spAppWebUrl);
                    clientContext.set_webRequestExecutorFactory(factory);
                    var appContextSite = new SP.AppContextSite(clientContext, hosturl);
                    web = appContextSite.get_web();
                    if (useApp)
                        clientContext.$7_0.get_webRequest()._headers['Authorization'] = "Bearer " + AppOnlyAccessTokenForSPAPP;
                }
                else {
                    clientContext = new SP.ClientContext(hosturl);
                    web = clientContext.get_web();
                }
            }
            else {
                clientContext = SP.ClientContext.get_current();
                web = clientContext.get_web();
            }

            return {
                ClientContext: clientContext,
                Web: web
            };
        };
    }

    SharePointFileService.$inject = ['$http', '$q'];
    function SharePointFileService($http, $q) {
        this.uploadItemAttachment = function (file, newName, attachmentURL) {
            var defer = $q.defer();

            var getFile = getFileBuffer();
            getFile.done(function (arrayBuffer) {

                // Add the file to the SharePoint folder.
                var addFile = addFileToFolder(arrayBuffer);
                addFile.done(function (file, status, xhr) {

                    // Get the list item that corresponds to the uploaded file.
                    //var getItem = getListItem(file.d.ListItemAllFields.__deferred.uri);
                    //getItem.done(function (listItem, status, xhr) {

                    //    // Change the display name and title of the list item.
                    //    var changeItem = updateListItem(listItem.d.__metadata);
                    //    changeItem.done(function (data, status, xhr) {
                    //        defer.resolve(data);
                    //    });
                    //    changeItem.fail(function (err) {
                    //        defer.reject(err);
                    //    });
                    //});
                    //getItem.fail(function (err) {
                    //    defer.reject(err);
                    //});
                    defer.resolve(file);
                });
                addFile.fail(function (err) {
                    defer.reject(err);
                });
            });
            getFile.fail(function (err) {
                defer.reject(err);
            });

            // Get the local file as an array buffer.
            function getFileBuffer() {
                var deferred = jQuery.Deferred();
                var reader = new FileReader();
                reader.onloadend = function (e) {
                    deferred.resolve(e.target.result);
                }
                reader.onerror = function (e) {
                    deferred.reject(e.target.error);
                }
                reader.readAsArrayBuffer(file);
                return deferred.promise();
            }

            // Add the file to the file collection in the Shared Documents folder.
            function addFileToFolder(arrayBuffer) {

                // Get the file name from the file input control on the page.
                //var parts = fileInput[0].value.split('\\');
                //var fileName = parts[parts.length - 1];
                //var fileName = file.name;
                // Construct the endpoint.
                var fileCollectionEndpoint = String.format(
                                attachmentURL +
                                "/add(FileName='{0}')",
                                newName);

                // Send the request and return the response.
                // This call returns the SharePoint file.
                return jQuery.ajax({
                    url: fileCollectionEndpoint,
                    type: "POST",
                    data: arrayBuffer,
                    processData: false,
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "X-RequestDigest": jQuery("#__REQUESTDIGEST").val(),
                        "content-length": arrayBuffer.byteLength
                    }
                });
            }

            // Get the list item that corresponds to the file by calling the file's ListItemAllFields property.
            function getListItem(fileListItemUri) {

                // Send the request and return the response.
                return jQuery.ajax({
                    url: fileListItemUri,
                    type: "GET",
                    headers: {
                        "accept": "application/json;odata=verbose"
                    }
                });
            }

            // Change the display name and title of the list item.
            function updateListItem(itemMetadata) {

                // Define the list item changes. Use the FileLeafRef property to change the display name. 
                // For simplicity, also use the name as the title. 
                // The example gets the list item type from the item's metadata, but you can also get it from the
                // ListItemEntityTypeFullName property of the list.
                var body = String.format("{{'__metadata':{{'type':'{0}'}},'FileLeafRef':'{1}','Title':'{2}'}}",
                    itemMetadata.type, newName, newName);

                // Send the request and return the promise.
                // This call does not return response content from the server.
                return jQuery.ajax({
                    url: itemMetadata.uri,
                    type: "POST",
                    data: body,
                    headers: {
                        "X-RequestDigest": jQuery("#__REQUESTDIGEST").val(),
                        "content-type": "application/json;odata=verbose",
                        "content-length": body.length,
                        "IF-MATCH": itemMetadata.etag,
                        "X-HTTP-Method": "MERGE"
                    }
                });
            }

            return defer.promise;
        }
    }

    SharePointJSOMService.$inject = ['$q', '$http', 'JSOMExecutor', 'HttpExecutor'];
    function SharePointJSOMService($q, $http, JSOMExecutor, HttpExecutor) {
        this.getCurrentUser = function () {
            var deferred = $q.defer();

            if (appMode) {
                var context = new SP.ClientContext(spAppWebUrl);
                var factory = new SP.ProxyWebRequestExecutorFactory(spAppWebUrl);
                context.set_webRequestExecutorFactory(factory);
                var appContextSite = new SP.AppContextSite(context, spAppWebUrl);
                //context.$7_0.get_webRequest()._headers['Authorization'] = "Bearer " + AppOnlyAccessTokenForSPAPP;
                var outweb = appContextSite.get_web();

                var user = outweb.get_currentUser();
                context.load(outweb, 'EffectiveBasePermissions');
                context.load(user);

                context.executeQueryAsync(
                    function (data) {
                        appContextSite = new SP.AppContextSite(context, hostweburl);
                        context.$7_0.get_webRequest()._headers['Authorization'] = "Bearer " + AppOnlyAccessTokenForSPAPP;
                        web = appContextSite.get_web();
                        user = web.ensureUser(user.get_loginName());
                        var permissions = outweb.get_effectiveBasePermissions();
                        context.load(user);

                        var userGroups = user.get_groups();
                        context.load(userGroups);

                        context.executeQueryAsync(
                            function (sender, args) {
                                deferred.resolve({ 'user': user, 'permissions': permissions, 'userGroups': userGroups });
                            },
                            function (sender, args) {
                                deferred.reject(args.get_message());
                            });
                    },
                    function (sender, args) {
                        deferred.reject(args.get_message());
                    });
            }
            else {
                var context = SP.ClientContext.get_current();
                var web = context.get_web();

                var user = web.get_currentUser();
                context.load(user);

                var userGroups = user.get_groups();
                context.load(userGroups);

                context.executeQueryAsync(
                    function (sender, args) {
                        deferred.resolve({ 'user': user, 'userGroups': userGroups });
                    },
                    function (sender, args) {
                        deferred.reject(args.get_message());
                    });
            }

            return deferred.promise;
        }

        this.getSiteUser = function (webUrl, userMail) {
            var deferred = $q.defer();

            if (appMode) {
                var context = new SP.ClientContext(spAppWebUrl);
                var factory = new SP.ProxyWebRequestExecutorFactory(spAppWebUrl);
                context.set_webRequestExecutorFactory(factory);
                var appContextSite = new SP.AppContextSite(context, spAppWebUrl);
                //context.$7_0.get_webRequest()._headers['Authorization'] = "Bearer " + AppOnlyAccessTokenForSPAPP;
                var web = appContextSite.get_web();

                var user = web.get_currentUser();
                context.load(user);

                context.executeQueryAsync(
                    function (data) {
                        appContextSite = new SP.AppContextSite(context, webUrl);
                        context.$7_0.get_webRequest()._headers['Authorization'] = "Bearer " + AppOnlyAccessTokenForSPAPP;
                        web = appContextSite.get_web();
                        user = web.ensureUser(user.get_loginName());
                        context.load(user);

                        context.executeQueryAsync(
                            function (sender, args) {
                                deferred.resolve({ 'user': user });
                            },
                            function (sender, args) {
                                deferred.reject(args.get_message());
                            });
                    },
                    function (sender, args) {
                        deferred.reject(args.get_message());
                    });
            }
            else {
                var context = new SP.ClientContext(webUrl);
                var web = context.get_web();

                var user = web.get_currentUser();
                context.load(user);
                context.executeQueryAsync(
                    function (sender, args) {
                        deferred.resolve({ 'user': user });
                    },
                    function (sender, args) {
                        deferred.reject(args.get_message());
                    });
            }

            return deferred.promise;
        }
        this.getCurrentSiteUrl = function () {
            var deferred = $q.defer();

            var executor = JSOMExecutor(hostweburl, true);

            executor.ClientContext.load(executor.Web);
            executor.ClientContext.executeQueryAsync(
                function (data) {
                    deferred.resolve(executor.Web.get_url());
                },
                function (err) {
                    deferred.reject();
                });

            return deferred.promise;
        };
        this.getGroupUsersByName = function (groupName, weburl) {
            var deferred = $q.defer();

            var context = SP.ClientContext.get_current();
            var excutor = JSOMExecutor(weburl, appMode);
            var collGroup = excutor.Web.get_siteGroups();
            var oGroup = collGroup.getByName(groupName);
            var users = oGroup.get_users();
            excutor.ClientContext.load(users);
            excutor.ClientContext.executeQueryAsync(
               function (data) {
                   deferred.resolve(users);
               },
               function (sender, args) {
                   deferred.reject(args);
               }
            );

            return deferred.promise;
        }
        this.getUserProfile = function (weburl) {
            var deferred = $q.defer();
            var excutor = JSOMExecutor(weburl, appMode);
            var userInfoList = excutor.Web.get_siteUserInfoList();

            var camlQuery = new SP.CamlQuery();

            camlQuery.set_viewXml('<View></View>');

            var collListItem = userInfoList.getItems(camlQuery);

            excutor.ClientContext.load(collListItem);
            excutor.ClientContext.executeQueryAsync(
                function (data) {
                    deferred.resolve(collListItem);
                },
                function (sender, args) {
                    deferred.reject(args);
                }
            );
            return deferred.promise;
        }
        this.getUserProfileById = function (userID) {
            var deferred = $q.defer();
            var context = SP.ClientContext.get_current();
            var userInfoList = context.get_web().get_siteUserInfoList();

            var camlQuery = new SP.CamlQuery();

            camlQuery.set_viewXml('<View><Query><Where><Eq><FieldRef Name=\'ID\'/>' + '<Value Type=\'Number\'>' + userID + '</Value></Eq>' +
            '</Where></Query><RowLimit>1</RowLimit></View>');

            var collListItem = userInfoList.getItems(camlQuery);

            context.load(collListItem);
            context.executeQueryAsync(
                function (data) {
                    deferred.resolve(collListItem.itemAt(0));
                },
                function (sender, args) {
                    deferred.reject(args);
                }
            );
            return deferred.promise;
        }
        this.updateUserTitleById = function (userID, title) {
            //var deferred = $q.defer();
            var context = SP.ClientContext.get_current();
            var userInfoList = context.get_web().get_siteUserInfoList();

            var camlQuery = new SP.CamlQuery();

            camlQuery.set_viewXml('<View><Query><Where><Eq><FieldRef Name=\'ID\'/>' + '<Value Type=\'Number\'>' + userID + '</Value></Eq>' +
            '</Where></Query><RowLimit>1</RowLimit></View>');

            var collListItem = userInfoList.getItems(camlQuery);
            context.load(collListItem);

            context.executeQueryAsync(
                function (data) {
                    var listItem = collListItem.itemAt(0)
                    listItem.set_item('Title', title);
                    listItem.update();
                    context.executeQueryAsync();
                    //deferred.resolve(collListItem.itemAt(0));
                },
                function (sender, args) {
                    //deferred.reject(args);
                }
            );
            //return deferred.promise;
        }
        this.getUserById = function (userID) {
            var deferred = $q.defer();

            var context = SP.ClientContext.get_current();
            var user = context.get_web().getUserById(userID);
            context.load(user);
            context.executeQueryAsync(
               function (data) {
                   deferred.resolve(user);
               },
               function (sender, args) {
                   deferred.reject(args);
               }
            );

            return deferred.promise;
        }
        this.getAttachments = function (listName, folderName, weburl, useApp) {
            var deferred = $q.defer();

            var attachments = [];

            var executor = JSOMExecutor(weburl, useApp);

            var list = executor.Web.get_lists().getByTitle(listName);
            var query = SP.CamlQuery.createAllItemsQuery();
            query.set_folderServerRelativeUrl(listName + '/' + folderName);
            var listItems = list.getItems(query);
            executor.ClientContext.load(listItems, 'Include(ID, Title, ContentType, File, CreatedBy)');
            executor.ClientContext.executeQueryAsync(
                function (sender, args) {
                    var listItemEnumerator = listItems.getEnumerator();
                    while (listItemEnumerator.moveNext()) {
                        var listItem = listItemEnumerator.get_current();

                        var createdBy = listItem.get_item('CreatedBy');

                        var file = listItem.get_file();
                        attachments.push({
                            ID: listItem.get_item('ID'),
                            Title: file.get_name(),
                            url: file.get_serverRelativeUrl(),
                            CreatedBy: {
                                ID: createdBy && createdBy.get_lookupId(),
                                Title: createdBy && createdBy.get_lookupValue()
                            }
                        })
                    }

                    deferred.resolve(attachments);
                },
                function (sender, args) {
                    deferred.reject(args.get_message());
                });

            return deferred.promise;
        };
        this.createFolder = function (listName, folderName, weburl, useApp) {
            var deferred = $q.defer();

            var executor = JSOMExecutor(weburl, useApp);

            var list = executor.Web.get_lists().getByTitle(listName);

            var itemCreateInfo = new SP.ListItemCreationInformation();
            itemCreateInfo.set_underlyingObjectType(SP.FileSystemObjectType.folder);
            itemCreateInfo.set_leafName(folderName);
            var listItem = list.addItem(itemCreateInfo);
            listItem.update();

            executor.ClientContext.load(listItem);
            executor.ClientContext.executeQueryAsync(
                function (sender, args) {
                    deferred.resolve();
                },
                function (sender, args) {
                    deferred.reject(args.get_message());
                });

            return deferred.promise;
        };
        this.upload = function (serverUrl, folder, item, userModel) {
            var defer = $q.defer();

            getFileBuffer().then(
                function (data) {
                    addFileToFolder(data, userModel).then(
                        function (response) {
                            var file = response.data;

                            item.url = file.d.ServerRelativeUrl;
                            item.GUID = file.d.UniqueId;

                            getListItem(file.d.ListItemAllFields.__deferred.uri).then(
                                function (response) {
                                    var listItem = response.data;

                                    item.ID = listItem.d.ID;

                                    updateListItem(listItem.d.__metadata).then(
                                        function (response) {
                                            defer.resolve();
                                        },
                                        function (response) {
                                            defer.reject();
                                        });
                                },
                                function (response) {
                                    defer.reject();
                                });
                        },
                        function (response) {
                            defer.reject();
                        });
                },
                function (err) {
                    defer.reject(err);
                });

            function getFileBuffer() {
                var deferred = $q.defer();

                var reader = new FileReader();
                reader.readAsArrayBuffer(item.fileInput);
                reader.onloadend = function (e) {
                    deferred.resolve(e.target.result);
                }
                reader.onerror = function (e) {
                    deferred.reject(e.target.error);
                }

                return deferred.promise;
            }

            function addFileToFolder(arrayBuffer, userModel) {
                var fileCollectionEndpoint = String.format("{0}/_api/web/getfolderbyserverrelativeurl('{1}')/files" + "/add(overwrite=true, url='{2}')", serverUrl, folder, item.Title);

                var headers = {
                    "accept": "application/json;odata=verbose",
                    "content-length": arrayBuffer.byteLength,
                    'X-RequestDigest': $("#__REQUESTDIGEST").val()
                };
                return HttpExecutor.postWithoutTransformRequest(fileCollectionEndpoint, arrayBuffer, headers, userModel);
            }

            function getListItem(fileListItemUri) {
                var headers = {
                    "accept": "application/json;odata=verbose",
                    'X-RequestDigest': $("#__REQUESTDIGEST").val()
                };
                return HttpExecutor.get(fileListItemUri, headers);
            }

            function updateListItem(itemMetadata) {
                var body = "{'__metadata':{'type':'" + itemMetadata.type + "'}," + JSON.stringify({ CreatedById: item.CreatedBy.ID }).slice(1, -1) + "}";

                var headers = {
                    "content-type": "application/json;odata=verbose",
                    "IF-MATCH": itemMetadata.etag,
                    "X-HTTP-Method": "MERGE",
                    'X-RequestDigest': $("#__REQUESTDIGEST").val()
                };
                return HttpExecutor.post(itemMetadata.uri, body, headers);
            }

            return defer.promise;
        }

        this.getChoiceFieldChoices = function (listName, fieldName, weburl, useApp) {
            var deferred = $q.defer();

            var executor = JSOMExecutor(weburl, useApp);

            var list = executor.Web.get_lists().getByTitle(listName);

            var field = executor.ClientContext.castTo(list.get_fields().getByInternalNameOrTitle(fieldName), SP.FieldChoice);

            executor.ClientContext.load(field);
            executor.ClientContext.executeQueryAsync(
                function (sender, args) {
                    deferred.resolve(field.get_choices());
                },
                function (sender, args) {
                    deferred.reject(args.get_message());
                });

            return deferred.promise;
        };

        this.query = function (listName, camlQueryString, pagingInfo, weburl, useApp, folderRelativeURL, exp, includeFields) {
            var deferred = $q.defer();

            var executor = JSOMExecutor(weburl, useApp);
            var list;
            if (/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(listName)) {
                list = executor.Web.get_lists().getById(listName);
            } else {
                list = executor.Web.get_lists().getByTitle(listName);
            }
            

            var camlQuery = new SP.CamlQuery();
            camlQuery.set_folderServerRelativeUrl(folderRelativeURL);
            camlQuery.set_viewXml(camlQueryString);

            if (pagingInfo != undefined && pagingInfo != '') {
                var position = new SP.ListItemCollectionPosition();
                position.set_pagingInfo(pagingInfo);
                camlQuery.set_listItemCollectionPosition(position);
            }

            var listItems = list.getItems(camlQuery);
            
            var listFields = list.get_fields();
            if (includeFields) {
                executor.ClientContext.load(listFields);
            }

            if (!exp) {
                executor.ClientContext.load(listItems);
            }
            else {
                executor.ClientContext.load(listItems, exp);
            }
            executor.ClientContext.executeQueryAsync(
                function (sender, args) {
                    if (includeFields)
                        deferred.resolve([listItems, listFields]);
                    else
                        deferred.resolve(listItems);
                },
                function (sender, args) {
                    deferred.reject(args.get_message());
                });

            return deferred.promise;
        };

        this.delete = function (listName, entityID, weburl, useApp) {
            var deferred = $q.defer();

            var executor = JSOMExecutor(weburl, useApp);

            var list = executor.Web.get_lists().getByTitle(listName);
            var listItem = list.getItemById(entityID);
            listItem.deleteObject();

            executor.ClientContext.executeQueryAsync(
                function (sender, args) {
                    deferred.resolve();
                },
                function (sender, args) {
                    deferred.reject(args.get_message());
                });

            return deferred.promise;
        };
    }

    SharePointResource.$inject = ['$resource', 'config'];
    function SharePointResource($resource, config) {

        // Initialize the QueryString to conditionally target the Host Url
        function getTargetQueryString(webUrl) {
            var targetQueryString;
            if (appMode)
                targetQueryString = "?@target=" + webUrl;

            return targetQueryString;
        }

        return {
            // Gets list items and creates list items.
            Items: function (ODataQuery, webUrl) {

                var queryString = getTargetQueryString(webUrl);

                if (queryString && ODataQuery) {
                    queryString = queryString + "&" + ODataQuery;
                }
                else if (ODataQuery) {
                    queryString = "?" + ODataQuery;
                }

                queryString = queryString || "";

                return $resource(webUrl + "/_api/web/lists/GetByTitle(':SPList')/items" + queryString,
                { SPList: '' },
                {
                    get: { method: 'GET', headers: { 'accept': 'application/json;odata=verbose', 'Authorization': "Bearer " + AppOnlyAccessTokenForSPHost } },
                    create: {
                        method: 'POST',
                        isArray: false,
                        headers: {
                            'accept': 'application/json;odata=verbose',
                            'X-RequestDigest': $("#__REQUESTDIGEST").val(),
                            'content-type': 'application/json;odata=verbose',
                            'Authorization': "Bearer " + AppOnlyAccessTokenForSPHost
                        }
                    }
                })
            },

            // Get, Update and Delete a single list item.
            Item: function (weburl) {

                var queryString = getTargetQueryString(weburl) || "";

                return $resource(weburl+"/_api/web/lists/GetByTitle(':SPList')/items(:SPListItem)" + queryString,
                { weburl: '', SPList: '', SPListItem: '' },
                {
                    get: { method: 'GET', headers: { 'accept': 'application/json;odata=verbose', 'Authorization': "Bearer " + AppOnlyAccessTokenForSPHost } },
                    update: {
                        method: 'POST',
                        headers: {
                            'accept': 'application/json;odata=verbose',
                            'X-RequestDigest': $("#__REQUESTDIGEST").val(),
                            'IF-MATCH': '*',
                            'X-HTTP-Method': 'MERGE',
                            'content-type': 'application/json;odata=verbose',
                            'Authorization': "Bearer " + AppOnlyAccessTokenForSPHost
                        }
                    },
                    delete: {
                        method: 'POST',
                        headers: {
                            'accept': 'application/json;odata=verbose',
                            'X-RequestDigest': $("#__REQUESTDIGEST").val(),
                            'IF-MATCH': '*',
                            'X-HTTP-Method': 'DELETE',
                            'content-type': 'application/json;odata=verbose',
                            'Authorization': "Bearer " + AppOnlyAccessTokenForSPHost
                        }
                    }
                })
            },

            // Get List object By List Name
            GetListByName: function (webUrl) {
                var queryString = getTargetQueryString(webUrl) || "";

                return $resource(webUrl + "/_api/web/lists/GetByTitle(':SPList')" + queryString,
                { SPList: '' },
                {
                    get: { method: 'GET', headers: { 'accept': 'application/json;odata=verbose', 'Authorization': "Bearer " + AppOnlyAccessTokenForSPHost } }
                });
            }
        };

        //return SharePointResource;
    }

    SharePointAPIService.$inject = ['$q', 'SharePointResource'];
    function SharePointAPIService($q, SharePointResource) {
        var spData = {

            // CRUD and Query operations for list items.
            spListItem: {
                Create: function (ListName, webUrl) {
                    var deferred = $q.defer();
                    SharePointResource.GetListByName(webUrl).get({ SPList: ListName }).$promise.then(function (data) {
                        var newItem = { '__metadata': { 'type': data.d.ListItemEntityTypeFullName } };
                        deferred.resolve(newItem);
                    },
                    function (data) {
                        deferred.reject(data);
                    });
                    //ListName = ListName.replace(" ", "_x0020_");
                    //return { '__metadata': { 'type': 'SP.Data.' + ListName + 'ListItem' } };
                    return deferred.promise;
                },
                Get: function (ListName, ID, webUrl) {
                    return SharePointResource.Item(webUrl).get({ weburl: webUrl, SPList: ListName, SPListItem: ID }).$promise;
                },
                Update: function (ListName, Item, webUrl) {
                    if (Item.ID) {
                        return SharePointResource.Item(webUrl).update({ SPList: ListName, SPListItem: Item.ID }, Item).$promise;
                    }
                    else
                        return SharePointResource.Items('', webUrl).create({ SPList: ListName }, Item).$promise;
                },
                Delete: function (ListName, ID, webUrl) { return SharePointResource.Item(webUrl).delete({ SPList: ListName, SPListItem: ID }, '').$promise; },
                Query: function (ListName, ODataQuery, webUrl) {
                    return SharePointResource.Items(ODataQuery, webUrl).get({ SPList: ListName, SPUrl: webUrl }).$promise;
                }
            }
        };

        return spData;
    }

    SharePointFormService.$inject = ['$q', '$http', 'SharePointAPIService'];
    function SharePointFormService($q, $http, SharePointAPIService){
        this.loadItem = function (listTitle, itemId, weburl) {
            return SharePointAPIService.spListItem.Get(listTitle, itemId, weburl);
        }

        this.saveItem = function (listTitle, formItem, formObject, weburl) {
            var deferred = $q.defer();
            SharePointAPIService.spListItem.Create(listTitle, weburl).then(
                function (item) {

                    for (var i = 0; i < formObject.length; i++) {
                        var modelName = formObject[i].name;
                        if (modelName) {
                            item[modelName] = formItem[modelName];
                        }
                    }

                    SharePointAPIService.spListItem.Update(listTitle, item, weburl).then(
                        function (data) {
                            deferred.resolve(data);
                        },
                        function (err) {
                            deferred.reject(err);
                        });
                },
                function (err) {
                    deferred.reject(err);
                });
            return deferred.promise;
        }

        this.loadMasterItems = function (listTitle, ODataQuery, weburl) {
            return SharePointAPIService.spListItem.Query(listTitle, ODataQuery, weburl);
        }
    }

    SharePointListService.$inject = ['$q', '$http', 'SharePointAPIService'];
    function SharePointListService($q, $http, SharePointAPIService){
        this.loadItems = function (listTitle, ODataQuery, weburl) {
            return SharePointAPIService.spListItem.Query(listTitle, ODataQuery, weburl);
        }

        this.saveItems = function (listTitle, listItems, weburl) {
            var mainDeferred = $q.defer();
            SharePointAPIService.spListItem.Create(listTitle, weburl).then(
                function (item) {
                    var promises = [];

                    for (var i = 0; i < listItems.length; i++) {
                        var deferred = $q.defer();

                        var listItem = listItems[i];
                        for (var p in listItem) {
                            item[p] = listItem[p];
                        }

                        promises.push(SharePointAPIService.spListItem.Update(listTitle, item));
                    }

                    $q.all(promises).then(
                        function (results) {
                            mainDeferred.resolve(results);
                        },
                        function (errs) {
                            mainDeferred.reject(errs);
                        });
                },
                function (err) {
                    mainDeferred.reject(err);
                });

            return mainDeferred.promise;
        }

        this.deleteItem = function (listTitle, ID, weburl) {
            var deferred = $q.defer();
            SharePointAPIService.spListItem.Delete(listTitle, ID, weburl).then(
            function (data) {
                deferred.resolve(data);
            },
            function (err) {
                deferred.reject(err);
            });
            return deferred.promise;
        }

    }

    SharePointPermissionsService.$inject = ['$q','JSOMExecutor']
    function SharePointPermissionsService($q, JSOMExecutor) {
        this.setListItemPermissionByUserId = function (userRoles, listTitle, itemId) {
            //var listTitle = 'Task';
            //var itemId = 1;
            //var userId = 8;
            //var roleType = SP.RoleType.administrator;//administrator,contributor,editor,guest,reader,webDesigner

            // var userRoles = [{ 'userId': 1, 'roleType': '' }, { 'userId': 2, 'roleType': '' }];
            //jsomWEBURL
            var jsomExec = new JSOMExecutor('https://serviceme.sharepoint.cn/sites/servicemedemo/eproject', true);
            
            var deferred = $q.defer();

            //var context = SP.ClientContext.get_current();
            //var listItem = context.get_web().get_lists().getByTitle(listTitle).getItemById(itemId);
            var listItem = jsomExec.Web.get_lists().getByTitle(listTitle).getItemById(itemId);

            listItem.breakRoleInheritance(false, true);

            for (var i = 0; i < userRoles.length; i++) {
                var userRole = userRoles[i];


                var user = jsomExec.Web.getUserById(userRole.userId);
                var roleDefBindingColl = SP.RoleDefinitionBindingCollection.newObject(jsomExec.ClientContext);
                roleDefBindingColl.add(jsomExec.Web.get_roleDefinitions().getByType(userRole.roleType));
                listItem.get_roleAssignments().add(user, roleDefBindingColl);


                //var user = context.get_web().getUserById(userRole.userId);
                //var roleDefBindingColl = SP.RoleDefinitionBindingCollection.newObject(context);
                //roleDefBindingColl.add(context.get_web().get_roleDefinitions().getByType(userRole.roleType));
                //listItem.get_roleAssignments().add(user, roleDefBindingColl);
            }

            jsomExec.ClientContext.executeQueryAsync(
            //context.executeQueryAsync(
               function (data) {
                   deferred.resolve(data);
               },
               function (sender, args) {
                   deferred.reject(args);
               }
            );

            return deferred.promise;
        }
        this.setListItemPermissionByGroupId = function (groupRoles, listTitle, itemId) {
            //var listTitle = 'Task';
            //var itemId = 1;
            //var userId = 8;
            //var roleType = SP.RoleType.administrator;//administrator,contributor,editor,guest,reader,webDesigner

            // var groupRoles = [{ 'groupId': 1, 'roleType': '' }, { 'groupId': 2, 'roleType': '' }];

            var deferred = $q.defer();
            //jsomWEBURL
            var jsomExec = new JSOMExecutor('https://serviceme.sharepoint.cn/sites/servicemedemo/eproject', true);

            //var context = SP.ClientContext.get_current();
            //var listItem = context.get_web().get_lists().getByTitle(listTitle).getItemById(itemId);

            var listItem = jsomExec.Web.get_lists().getByTitle(listTitle).getItemById(itemId);

            listItem.breakRoleInheritance(false, true);

            for (var i = 0; i < groupRoles.length; i++) {
                var groupRole = groupRoles[i];

                //var user = context.get_web().getUserById(groupRole.userId);
                var collGroup = jsomExec.Web.get_siteGroups();
                var oGroup = collGroup.getById(groupRole.groupId);
                var roleDefBindingColl = SP.RoleDefinitionBindingCollection.newObject(jsomExec.ClientContext);
                roleDefBindingColl.add(jsomExec.Web.get_roleDefinitions().getByType(groupRole.roleType));
                listItem.get_roleAssignments().add(oGroup, roleDefBindingColl);

                //var collGroup = context.get_web().get_siteGroups();
                //var oGroup = collGroup.getById(groupRole.groupId);
                //var roleDefBindingColl = SP.RoleDefinitionBindingCollection.newObject(context);
                //roleDefBindingColl.add(context.get_web().get_roleDefinitions().getByType(groupRole.roleType));
                //listItem.get_roleAssignments().add(oGroup, roleDefBindingColl);
            }

            jsomExec.ClientContext.executeQueryAsync(
               function (data) {
                   deferred.resolve(data);
               },
               function (sender, args) {
                   deferred.reject(args);
               }
            );

            return deferred.promise;
        }
        this.deleteListItemPermissionByUserId = function (userRoles, listTitle, itemId) {
            var deferred = $q.defer();

            //var context = SP.ClientContext.get_current();
            //var listItem = context.get_web().get_lists().getByTitle(listTitle).getItemById(itemId);
            //jsomWEBURL
            var jsomExec = new JSOMExecutor('https://serviceme.sharepoint.cn/sites/servicemedemo/eproject', true);

            for (var i = 0; i < userRoles.length; i++) {
                var userRole = userRoles[i];
                var user = jsomExec.Web.getUserById(userRole.userId);
                listItem.get_roleAssignments().getByPrincipal(user).deleteObject();

                //var user = context.get_web().getUserById(userRole.userId);
                //listItem.get_roleAssignments().getByPrincipal(user).deleteObject();
            }

            jsomExec.ClientContext.executeQueryAsync(
               function () {
                   deferred.resolve(user);
               },
               function (sender, args) {
                   deferred.resolve(args);
               }
            );

            return deferred.promise;
        }
    }
})();
