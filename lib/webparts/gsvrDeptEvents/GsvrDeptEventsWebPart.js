"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_1 = require("@pnp/sp");
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var sp_http_1 = require("@microsoft/sp-http");
var sp_lodash_subset_1 = require("@microsoft/sp-lodash-subset");
var GsvrDeptEventsWebPart_module_scss_1 = require("./GsvrDeptEventsWebPart.module.scss");
var strings = require("GsvrDeptEventsWebPartStrings");
//global vars
var userDept = "";
var GsvrDeptEventsWebPart = (function (_super) {
    __extends(GsvrDeptEventsWebPart, _super);
    function GsvrDeptEventsWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        // get all the user properties
        _this.getuser = new Promise(function (resolve, reject) {
            // SharePoint PnP Rest Call to get the User Profile Properties
            return sp_1.sp.profiles.myProperties.get().then(function (result) {
                var props = result.UserProfileProperties;
                var propValue = "";
                var userDepartment = "";
                props.forEach(function (prop) {
                    //this call returns key/value pairs so we need to look for the Dept Key
                    if (prop.Key == "Department") {
                        // set our global var for the users Dept.
                        userDept += prop.Value;
                    }
                });
                return result;
            }).then(function (result) {
                _this._getListData().then(function (response) {
                    _this._renderList(response.value);
                });
            });
        });
        return _this;
    }
    GsvrDeptEventsWebPart.prototype.render = function () {
        this.domElement.innerHTML = "\n      <div class=\"" + GsvrDeptEventsWebPart_module_scss_1.default.gsvrDeptEvents + "\">\n        <div class=\"" + GsvrDeptEventsWebPart_module_scss_1.default.container + "\">\n          <div class=\"" + GsvrDeptEventsWebPart_module_scss_1.default.row + "\">\n            <div class=\"" + GsvrDeptEventsWebPart_module_scss_1.default.column + "\">\n              <span class=\"" + GsvrDeptEventsWebPart_module_scss_1.default.title + "\">Welcome to SharePoint!</span>\n              <p class=\"" + GsvrDeptEventsWebPart_module_scss_1.default.subTitle + "\">Customize SharePoint experiences using Web Parts.</p>\n              <p class=\"" + GsvrDeptEventsWebPart_module_scss_1.default.description + "\">" + sp_lodash_subset_1.escape(this.properties.description) + "</p>\n              <a href=\"https://aka.ms/spfx\" class=\"" + GsvrDeptEventsWebPart_module_scss_1.default.button + "\">\n                <span class=\"" + GsvrDeptEventsWebPart_module_scss_1.default.label + "\">Learn more</span>\n              </a>\n              <h1>Events</h1>\n            <h3><div id=\"Events\"/></h3>\n            </div>\n          </div>\n        </div>\n      </div>";
    };
    Object.defineProperty(GsvrDeptEventsWebPart.prototype, "dataVersion", {
        get: function () {
            return sp_core_library_1.Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    // main REST Call to the list...passing in the deaprtment into the call to 
    //return a single list item
    GsvrDeptEventsWebPart.prototype._getListData = function () {
        return this.context.spHttpClient.get("https://girlscoutsrv.sharepoint.com/_api/web/lists/GetByTitle('TeamDashboardSettings')/Items?$filter=Title eq '" + userDept + "'", sp_http_1.SPHttpClient.configurations.v1)
            .then(function (response) {
            return response.json();
        });
    };
    GsvrDeptEventsWebPart.prototype._renderList = function (items) {
        var html = '';
        var libHTML = '';
        var siteURL = "";
        //list name
        var eventsListName = "";
        // items in the list
        var eventsItems = "";
        var date = new Date();
        var strToday = "";
        var mm = date.getMonth() + 1;
        console.log(mm);
        var dd = date.getDate();
        console.log(dd);
        var yyyy = date.getFullYear();
        console.log(yyyy);
        if (dd < 10) {
            dd = 0 + dd;
            console.log(dd);
        }
        if (mm < 10) {
            mm = 0 + mm;
            console.log(mm);
        }
        strToday = mm + "/" + dd + "/" + yyyy;
        console.log(strToday);
        items.forEach(function (item) {
            siteURL = item.DeptURL;
            eventsListName = item.CalURL;
        });
        //1st we need to override the current web to go to the department sites web
        var w = new sp_1.Web("https://girlscoutsrv.sharepoint.com" + siteURL);
        // then use PnP to query the list
        // CASIE IF YOU NEED MORE THAN 5 EVENTS JUST UPDATE THE NUMBER BELOW
        w.lists.getByTitle(eventsListName).items.filter("EventDate ge '" + strToday + "'").top(5)
            .get()
            .then(function (data) {
            console.log(data);
            for (var x = 0; x < data.length; x++) {
                //console.log(data[x].URL);
                //Title of the event
                console.log(data[x].Title);
                //Start Date - End Date
                console.log(data[x].EventDate + " - " + data[x].EndDate);
                //location of the event IF YOU NEED IT
                console.log(data[x].Location);
                //DESCRIPTION of the event IF YOU NEED IT
                console.log(data[x].Description);
                var titleLinkExample = "https://girlscoutsrv.sharepoint.com" + siteURL + "/Lists/" + eventsListName + "/DispForm.aspx?ID=" + data[x].Id;
                eventsItems += data[x].Title + "(" + data[x].EventDate + " - " + data[x].EndDate + ")" + '\r\n';
                // libHTML += `<p>${hrItems.toString()}</p>`;
            }
            document.getElementById("Events").innerText = eventsItems;
        }).catch(function (e) { console.error(e); });
        var listContainer = this.domElement.querySelector('#ListItems');
        listContainer.innerHTML = html;
    };
    // this is required to use the SharePoint PnP shorthand REST CALLS
    GsvrDeptEventsWebPart.prototype.onInit = function () {
        var _this = this;
        return _super.prototype.onInit.call(this).then(function (_) {
            sp_1.sp.setup({
                spfxContext: _this.context
            });
        });
    };
    GsvrDeptEventsWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return GsvrDeptEventsWebPart;
}(sp_webpart_base_1.BaseClientSideWebPart));
exports.default = GsvrDeptEventsWebPart;

//# sourceMappingURL=GsvrDeptEventsWebPart.js.map
