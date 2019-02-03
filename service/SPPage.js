
var PageManager = function (pernumber, maxsize, loadData) {
    return {
        getPageData: function () {
            if (Math.round(this.bigTotalItems / this.pernumber) <= this.bigCurrentPage &&
                this.bigCurrentPage % this.maxsize == 1 && this.bigCurrentPage > 1) {
                return false;
            }
            else
                return this.messages.slice(this.bigCurrentPage * this.pernumber - this.pernumber, this.bigCurrentPage * this.pernumber);
        },
        pushData: function (data) {
            if (data.length == 0) {
                this.remoteHasData = false;
            }
            else {
                this.messages = this.messages.concat(data);
                this.bigTotalItems = this.messages.length;
                if (data.length == this.getRowLimit()) {
                    this.messages.pop();
                }
            }
        },
        getLastItem: function () {
            var lastItem = undefined;
            if (this.messages != undefined && this.messages.length > 0)
                lastItem = this.messages[this.messages.length - 1];
            return lastItem;
        },
        getRowLimit: function () {
            return this.pernumber * this.maxsize + 1;
        },
        getPageInfo: function (info) {
            var pageInfo = '';
            for (var key in info) {
                pageInfo = pageInfo + '&p_' + key + '=' + info[key];
            }
            pageInfo = "Paged=TRUE" + pageInfo;
            return pageInfo;
        },
        showPageItems: function () {
            var data = this.getPageData();
            if (data == false && this.loadData != undefined
                && this.remoteHasData != false) {
                this.loadData();
            }
            else
                this.currentPageData = data || [];
        },
        reloadData: function () {
            this.messages = [];
            this.bigTotalItems = 0;
            this.bigCurrentPage = 1;
            this.remoteHasData = true;
            this.showPageItems();
        },
        reset: function () {
            this.messages = [];
            this.bigTotalItems = 0;
            this.bigCurrentPage = 1;
            this.remoteHasData = true;
        },
        remoteHasData: true,
        pernumber: pernumber,
        maxsize: maxsize,
        loadData: loadData,
        messages: [],
        bigTotalItems: 0,
        bigCurrentPage: 1
    };
};
