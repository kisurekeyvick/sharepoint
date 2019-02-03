/*
    caml语句
    可以使用谷歌商店下载caml插件
*/
const camlQueryDemo = function (currentUserID, conditions, pageInfo, rowLimit) {
    var deferred = $q.defer();

    var camlQuery = new CamlBuilder().View().RowLimit(rowLimit).Query();

    if (keyword) {
        camlQuery = camlQuery.And().TextField("Month").Contains(keyword);
        camlQuery = camlQuery.Or().TextField("Telephone").Contains(keyword);
    }

    if ($scope.currentUserInfo.CardId) {
        // 文本字段的值为
        camlQuery = camlQuery.And().TextField('EmployeeID').EqualTo($scope.currentUserInfo.CardId);
    }

    if (firstDay) {
        // 时间范围
        camlQuery = camlQuery.And().DateField("WeeklyStartDate").GreaterThanOrEqualTo(firstDay);
    }

    if (lastDay) {
        // 时间范围
        camlQuery = camlQuery.And().DateField("WeeklyEndDate").LessThanOrEqualTo(lastDay);
    }

    if (conditions) {
        // 多用户，查询其中的id为某个值
        camlQuery = camlQuery.Or().UserMultiField('ShareTo').IncludesSuchItemThat().Id().EqualTo(conditions.ID);
    }

    if (conditions) {
        // 用户字段的值在某个集合中
        camlQuery = camlQuery.And().UserField("CreatedBy").Id().In(conditions.userList);
    }

    if (conditions) {
        // 用户字段的文本值包含了某个关键字
        camlQuery = camlQuery.Or().UserField("AssignedTo").ValueAsText().Contains(conditions.Keyword);
    }

    if (a) {
        // OrderByDesc逆序，建议放在最后
        camlQuery = camlQuery.OrderByDesc('Month');
    }
    else {
        // 先逆序再正序
        camlQuery = camlQuery.OrderByDesc('WeeklyStartDate').ThenByDesc('ID');
    }

    if (a) {//any 括号里面的条件都可以满足 相当于"或"
        camlQuery = camlQuery.And().Any(
            CamlBuilder.Expression().UserMultiField('AssignedTo').IncludesSuchItemThat().Id().EqualTo(currentUserID).And().TextField('Status').EqualTo('Doing'),
            CamlBuilder.Expression().UserMultiField('Approver').IncludesSuchItemThat().Id().EqualTo(currentUserID).And().TextField('Status').EqualTo('Approving'))
    }

    if (a) {
        // 某个boolean值是否为true
        camlQuery = camlQuery.And().BooleanField('Assigned').IsTrue();
    }

    if (a) {//查询member用户组中存在我id的
        camlQuery = camlQuery.And().UserMultiField('Member').IncludesSuchItemThat().Id().EqualTo(currentUserID)
    }

    if (a) {
        var camlQuery = new CamlBuilder().View().RowLimit(rowLimit).Query().Where().CounterField('ID').GreaterThan(0);
    }

    if (a) {
        // 某个关联其他表的查询字段，它的文本值为
        camlQuery = camlQuery.And().LookupField('OrgUnit_x003a_OrgUnitID').ValueAsText().EqualTo(departmentCode);
        camlQuery = camlQuery.And().LookupField('OrgUnit_x003a_OrgUnitID').ValueAsText().BeginsWith(departmentCode);
    }

    if(d){
        // 时间范围  这里用了moment.js
        camlQuery = camlQuery.And().DateTimeField('MeetingStartDate').LessThanOrEqualTo(moment(conditions.MeetingDate).endOf('day').format('YYYY-MM-DDTHH:mm:ss'))
        .And().DateTimeField('MeetingEndDate').GreaterThanOrEqualTo(moment(conditions.MeetingDate).startOf('day').format('YYYY-MM-DDTHH:mm:ss'));
    }
}
