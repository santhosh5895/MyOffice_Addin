const domain = "https://souldev.vuram.com/suite/webapi/"
const apiKey = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJjMGI0MzJmZi04NDM3LTQwNTktOTY5OC0zOTc0Y2M1Y2JlMWEifQ.LmPQhX-osrinS5HNVIwMLrtMEX81Sg_ynYcxXa-5bkw";

export const dropdownOptions = [{
    name: 'Case Id',
    value: 'caseId',
    checked: true
},
{
    name: 'Subject',
    value: 'subject',
    checked: false
},
{
    name: 'Description',
    value: 'description',
    checked: false
},
{
    name: 'Reported By',
    value: 'reportedBy',
    checked: false
},
{
    name: 'Category',
    value: 'category',
    checked: false
}];
// export var selectedFilters = [];
export const dropdownSettings = {
    texts: {
        selectAll: "Select all filters",
        unselectAll: "Unselect all filters"
    },
    selectAll: true,
    minHeight: 100
};

export function http_get(to, data, callback) {
    var url = domain + to;
    $.ajax({
        url: url,
        data: data,
        success: callback,
        error: function (error) {
            console.log(error);
        },
        headers: {
            "Appian-API-Key": apiKey,
            "Access-Control-Allow-Origin": "*"
        },
        method: "GET"
    });
}