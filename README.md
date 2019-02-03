# 最近在整理这几年做过的项目经验，这一部分是基于sharepoint平台，使用的是angular.js
- 希望能够帮助到sharepoint开发者，如果好评，请不要吝啬给个star

## SPService.js
- SP API的集合

## SPPage.js
- SP 的分页

## 使用方法 - Demo

### `Js部分` - 需要注入SPService.js

```js
(function () {
    'use strict';

    angular.module('NgSP.Page')
    .controller('PageDemoController', PageDemoController);

    PageDemoController.$inject = ['$scope', '$q', 'SharePointJSOMService'];

    function PageDemoController($scope, $q, SharePointJSOMService) {
        function init() {
            $scope.pageInfo = new PageManager($scope.pernumber, $scope.maxsize, loadData);
        }

        function loadDate() {
            let  pageInfo = '';
            const last = $scope.pageInfo.getLastItem();
            const limit = $scope.pageInfo.getRowLimit();

            if (last) {
                const info = {
                    ID: last.ID || '0',
                    Date: last.Date.toISOString() || ''
                };

                pageInfo = $scope.pageInfo.getPageInfo(info);
            }

            let camlQuery = new CamlBuilder().View().RowLimit(rowLimit).Query().Where().CounterField('ID').GreaterThan(0);

            query('这是一个sharepoint列表名字',pageInfo, camlQuery, '子站点url').then((data) => {
                $scope.pageInfo.pushData(data);
                $scope.pageInfo.showPageItems();
            });        
        }

        function query(listName, pageInfo, camlQuery) {
            let deferred = $q.defer();
            
            SharePointJSOMService.query(listName, camlQuery.ToString(), pageInfo, url, true).then((data) => {
                const dataEnumerator = data.getEnumerator();
                const list = [];
                while (dataEnumerator.moveNext()) {
                    var listItem = dataEnumerator.get_current();
                    list.push(Colleague.createFrom(listItem));

                    deferred.resolve(list);
                },
                function (err) {
                    deferred.reject(err);
                });
            })

            return deferred.promise;
        }

        init();
        loadDate();
    }
})();
```
* query这个方法建议写在factory中，不知道factory的，[链接1](https://juejin.im/entry/56e786027db2a20052dc7356)
* camlQuery是sharepoint的查询，不知道的，[链接1](https://www.cnblogs.com/johnsonwong/archive/2011/02/27/1966008.html) 
[链接2](https://www.cnblogs.com/carysun/archive/2011/01/12/moss-caml.html),
你也可以下载CAML Query开发工具(打开sharepoint online的列表即可使用)
caml查询，我之前也做了一些笔记
```html
<div>
    <table>
        <thead>
            <tr>
                <th>标题</th>
                <th>工号</th>
            </tr>
        </thead>
        <tbody>
            <tr ng-repeat="data in pageInfo.currentPageData">
                <td ng-bind="data.Title"></td>
                <td ng-bind="data.JobNumber"></td>
            </tr>
        </tbody>
    </table>
    <div>
        <uib-pagination total-items="pageInfo.bigTotalItems" max-size="pageInfo.maxsize" ng-model="pageInfo.bigCurrentPage" previous-text="&lt;" next-text="&gt;" items-per-page="pageInfo.pernumber" boundary-link-numbers="true" rotate="false" ng-change="pageInfo.showPageItems()"></uib-pagination>
    </div>
</div>
```
* uib-pagination是一个插件，你可以忽略